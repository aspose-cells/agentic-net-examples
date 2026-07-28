// Title: Copy Selected Rows to a New Workbook with LightCells – Aspose.Cells C# Example
// Description: Loads a source workbook, copies specific rows (e.g., 0, 2, 4) to a destination worksheet using Cells.CopyRow, and saves the result with OoxmlSaveOptions and a custom WorksheetLightCellsDataProvider. The LightCells pipeline streams cell values during save, keeping memory usage minimal for large Excel files.
// Keywords: Aspose.Cells | C# | .NET | LightCells | WorksheetLightCellsDataProvider | copy rows | Cells.CopyRow | memory efficient Excel export | large workbook streaming | OoxmlSaveOptions
// Common Searches: Aspose.Cells copy specific rows C# | How to use LightCells to save a workbook with low memory | WorksheetLightCellsDataProvider example | Copy rows to new workbook without loading entire file | C# Aspose.Cells streaming export
// Developer Intent: Extract chosen rows from a source sheet and write them to a new workbook while minimizing RAM consumption.
// Use Cases: Create a lightweight report that contains only header and a few data rows from a massive Excel file. | Generate a filtered workbook for downstream processing without loading the full source into memory. | Build a subset workbook for archiving or sharing, using LightCells to stream data during save.
// AI Prompts: Show C# code that copies an array of row indices from one worksheet to another with Aspose.Cells and saves using LightCells for low memory usage. | Explain how to modify WorksheetLightCellsDataProvider to skip empty rows when streaming a workbook. | Adapt the example to copy selected columns instead of rows while still using LightCells for efficient saving.

using System;
using Aspose.Cells;

namespace LightCellsRowCopyDemo
{
    // Custom LightCellsDataProvider that streams data from a worksheet while saving
    // Loads a source workbook, copies specific rows (e.g., 0, 2, 4) to a destination worksheet using Cells.CopyRow, and saves the result with OoxmlSaveOptions and a custom WorksheetLightCellsDataProvider. The LightCells pipeline streams cell values during save, keeping memory usage minimal for large Excel files.
    public class WorksheetLightCellsDataProvider : LightCellsDataProvider
    {
        private readonly Worksheet _worksheet;
        private readonly int _maxRow;
        private readonly int _maxCol;
        private int _currentRow = -1;
        private int _currentCol = -1;

        public WorksheetLightCellsDataProvider(Worksheet worksheet)
        {
            _worksheet = worksheet;
            // Determine the last row/column that contain data
            _maxRow = worksheet.Cells.MaxDataRow;
            _maxCol = worksheet.Cells.MaxDataColumn;
        }

        // Called once for each sheet during saving
        public bool StartSheet(int sheetIndex)
        {
            // This provider only handles the first (and only) sheet
            return sheetIndex == 0;
        }

        // Return the next row index to be saved, or -1 to finish
        public int NextRow()
        {
            _currentRow++;
            _currentCol = -1; // reset column index for the new row
            return _currentRow <= _maxRow ? _currentRow : -1;
        }

        // Called after NextRow() returns a valid index; can modify row properties here
        public void StartRow(Row row)
        {
            // No special row handling required
        }

        // Return the next column index within the current row, or -1 to finish the row
        public int NextCell()
        {
            _currentCol++;
            return _currentCol <= _maxCol ? _currentCol : -1;
        }

        // Called after NextCell() returns a valid index; set the cell value
        public void StartCell(Cell cell)
        {
            Cell srcCell = _worksheet.Cells[_currentRow, _currentCol];
            if (srcCell != null && srcCell.Type != CellValueType.IsNull)
            {
                cell.PutValue(srcCell.Value);
            }
        }

        // Indicates whether string values should be gathered into a global string pool
        public bool IsGatherString()
        {
            return true;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook sourceWorkbook = new Workbook("source.xlsx");
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Create a new workbook that will hold the selected rows
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Define which rows to copy (example: rows 0, 2 and 4)
            int[] rowsToCopy = { 0, 2, 4 };
            int destRowIndex = 0;

            // Copy each selected row individually
            foreach (int srcRowIndex in rowsToCopy)
            {
                destSheet.Cells.CopyRow(sourceSheet.Cells, srcRowIndex, destRowIndex);
                destRowIndex++;
            }

            // Save the destination workbook using LightCells to reduce memory usage
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new WorksheetLightCellsDataProvider(destSheet)
            };

            destWorkbook.Save("selectedRows.xlsx", saveOptions);
        }
    }
}
