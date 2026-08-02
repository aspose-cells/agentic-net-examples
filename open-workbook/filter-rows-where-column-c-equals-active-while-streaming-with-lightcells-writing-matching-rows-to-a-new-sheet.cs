// Title: C# – Filter rows where column C = "Active" with Aspose.Cells LightCells streaming
// Description: Loads or creates a source workbook, then uses a custom LightCellsDataProvider to stream only rows whose column C (index 2) contains "Active" (including the header) into a new workbook named FilteredResult.xlsx on a sheet called "Filtered". This approach avoids loading the entire source file into memory.
// Keywords: Aspose.Cells LightCells filter rows | C# LightCells streaming | filter Excel rows by column value | custom LightCellsDataProvider example | write filtered data to new workbook | .NET Excel streaming | memory‑efficient Excel export
// Common Searches: Aspose.Cells filter rows while saving with LightCells | C# stream only rows where column equals a value using LightCellsDataProvider | How to export selected rows to a new Excel file without loading whole workbook | LightCells filter rows by criteria .NET | Create a lightweight Excel report with Aspose.Cells
// Developer Intent: Generate a new Excel file that contains only the rows from the source sheet where column C equals "Active", using LightCells streaming to keep memory usage low.
// Use Cases: Produce a compact report of active records from a massive dataset without fully loading the source file. | Export active customers to a separate workbook for downstream processing while preserving original row heights. | Create a filtered snapshot of a master sheet for archiving, minimizing RAM consumption.
// AI Prompts: Show a LightCellsDataProvider that filters rows based on a numeric threshold in column B. | Demonstrate how to copy cell styles and formatting in addition to values when using LightCells streaming. | Explain handling multiple worksheets with different filter criteria using LightCellsDataProvider.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace LightCellsFilterExample
{
    // Loads or creates a source workbook, then uses a custom LightCellsDataProvider to stream only rows whose column C (index 2) contains "Active" (including the header) into a new workbook named FilteredResult.xlsx on a sheet called "Filtered". This approach avoids loading the entire source file into memory.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "SourceData.xlsx";
                Workbook sourceWorkbook;

                // Load existing source workbook or create a sample one if it does not exist
                if (File.Exists(sourcePath))
                {
                    sourceWorkbook = new Workbook(sourcePath);
                }
                else
                {
                    sourceWorkbook = new Workbook();
                    Worksheet ws = sourceWorkbook.Worksheets[0];
                    ws.Cells[0, 0].PutValue("ID");
                    ws.Cells[0, 1].PutValue("Name");
                    ws.Cells[0, 2].PutValue("Status");
                    ws.Cells[1, 0].PutValue(1);
                    ws.Cells[1, 1].PutValue("Alice");
                    ws.Cells[1, 2].PutValue("Active");
                    ws.Cells[2, 0].PutValue(2);
                    ws.Cells[2, 1].PutValue("Bob");
                    ws.Cells[2, 2].PutValue("Inactive");
                    sourceWorkbook.Save(sourcePath);
                }

                Worksheet sourceSheet = sourceWorkbook.Worksheets[0]; // assume data is in the first sheet

                // Create a new workbook that will contain only the filtered rows
                Workbook filteredWorkbook = new Workbook();
                filteredWorkbook.Worksheets.Clear();               // start with a clean workbook
                filteredWorkbook.Worksheets.Add();                 // add a single worksheet for filtered data
                Worksheet targetSheet = filteredWorkbook.Worksheets[0];
                targetSheet.Name = "Filtered";

                // Prepare LightCells save options with a custom data provider
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    LightCellsDataProvider = new FilteredLightCellsDataProvider(sourceSheet)
                };

                // Save the filtered workbook using LightCells streaming
                filteredWorkbook.Save("FilteredResult.xlsx", saveOptions);
                Console.WriteLine("Filtered workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Custom LightCellsDataProvider that streams only rows where column C (index 2) equals "Active"
    class FilteredLightCellsDataProvider : LightCellsDataProvider
    {
        private readonly Worksheet _sourceSheet;
        private readonly List<int> _matchingRows = new List<int>();
        private int _maxColumn;               // number of columns to copy (0‑based count)
        private int _rowPointer = -1;         // index in _matchingRows
        private int _currentRowIndex = -1;    // actual row index in source sheet
        private int _colPointer = -1;         // column index while iterating cells

        public FilteredLightCellsDataProvider(Worksheet sourceSheet)
        {
            _sourceSheet = sourceSheet;
        }

        // Called when the saving process starts a worksheet.
        // We only provide data for the first (and only) sheet in the destination workbook.
        public bool StartSheet(int sheetIndex)
        {
            if (sheetIndex != 0) return false; // let Aspose.Cells handle any other sheets normally

            // Determine the range of columns to copy (up to the last used column)
            _maxColumn = _sourceSheet.Cells.MaxDataColumn + 1; // +1 because MaxDataColumn is zero‑based

            // Build the list of rows to copy:
            // - Always include the header row (row 0)
            // - Include rows where column C (index 2) has the value "Active"
            _matchingRows.Add(0); // header

            int lastDataRow = _sourceSheet.Cells.MaxDataRow; // will be -1 if sheet is empty
            if (lastDataRow >= 1)
            {
                for (int r = 1; r <= lastDataRow; r++)
                {
                    Cell criteriaCell = _sourceSheet.Cells[r, 2]; // column C
                    if (criteriaCell != null && criteriaCell.Type == CellValueType.IsString &&
                        string.Equals(criteriaCell.StringValue, "Active", StringComparison.OrdinalIgnoreCase))
                    {
                        _matchingRows.Add(r);
                    }
                }
            }

            // Reset iteration pointers
            _rowPointer = -1;
            _colPointer = -1;
            return true; // we will supply data for this sheet
        }

        // Returns the next row index to be written, or -1 when done.
        public int NextRow()
        {
            _rowPointer++;
            if (_rowPointer >= _matchingRows.Count)
                return -1; // no more rows

            _currentRowIndex = _matchingRows[_rowPointer];
            _colPointer = -1; // reset column pointer for the new row
            return _currentRowIndex; // row index in the destination sheet (zero‑based)
        }

        // Allows us to copy row height from the source sheet.
        public void StartRow(Row row)
        {
            Row srcRow = _sourceSheet.Cells.Rows[_currentRowIndex];
            if (srcRow != null)
                row.Height = srcRow.Height;
        }

        // Returns the next column index to be written, or -1 when the row is finished.
        public int NextCell()
        {
            _colPointer++;
            if (_colPointer >= _maxColumn)
                return -1; // end of columns for this row

            return _colPointer; // column index in the destination sheet
        }

        // Writes the cell value from the source sheet to the destination cell.
        public void StartCell(Cell cell)
        {
            Cell srcCell = _sourceSheet.Cells[_currentRowIndex, _colPointer];
            if (srcCell != null)
                cell.PutValue(srcCell.Value);
        }

        // Indicates whether the provider wants to gather string values for shared strings.
        // Returning false lets Aspose.Cells handle strings normally.
        public bool IsGatherString()
        {
            return false;
        }
    }
}
