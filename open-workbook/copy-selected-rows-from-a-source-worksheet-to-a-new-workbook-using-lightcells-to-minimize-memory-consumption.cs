// Title: C# – Copy Selected Rows to a New Workbook with Aspose.Cells LightCells (Low‑Memory)
// Description: Demonstrates how to load a source Excel file with a LightCellsDataHandler, collect rows 1‑5 (or any range), and stream those rows into a fresh workbook using a LightCellsDataProvider and OoxmlSaveOptions. The approach keeps memory usage minimal by processing rows on the fly instead of loading the entire source workbook.
// Keywords: Aspose.Cells | LightCells | C# | copy rows | low memory Excel processing | LightCellsDataHandler | LightCellsDataProvider | OoxmlSaveOptions | .NET Excel | row extraction | large workbook
// Common Searches: Aspose.Cells LightCells copy rows C# | how to extract specific rows from Excel with low memory | copy rows 1‑5 to new workbook using Aspose.Cells | LightCellsDataHandler example for row selection | stream selected rows to a new Excel file in .NET
// Developer Intent: Select a subset of rows from a source workbook and write them to a new workbook while keeping RAM consumption low.
// Use Cases: Create a lightweight summary file from a massive spreadsheet by extracting only the needed rows. | Generate a report workbook that contains filtered rows without loading the full source into memory. | Prepare a data slice for downstream processing or API response using Aspose.Cells streaming capabilities.
// AI Prompts: Show a C# snippet that copies rows 10‑20 from an Excel file to a new workbook using Aspose.Cells LightCells with minimal memory usage. | Explain how to modify the RowCollector to filter rows based on a cell value instead of a fixed index range. | Provide code to stream selected rows directly to an HTTP response using LightCellsDataProvider, avoiding temporary files.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace LightCellsRowCopyDemo
{
    // Collects selected rows from a source worksheet using LightCellsDataHandler
    // Demonstrates how to load a source Excel file with a LightCellsDataHandler, collect rows 1‑5 (or any range), and stream those rows into a fresh workbook using a LightCellsDataProvider and OoxmlSaveOptions. The approach keeps memory usage minimal by processing rows on the fly instead of loading the entire source workbook.
    class RowCollector : LightCellsDataHandler
    {
        // Rows that satisfy the selection criteria; each row is stored as an object array of cell values
        public List<object[]> SelectedRows { get; } = new List<object[]>();

        // Maximum number of columns among the selected rows (used later by the data provider)
        public int MaxColumnCount { get; private set; } = 0;

        // Define the range of rows to copy (example: rows 1 to 5, zero‑based indices)
        private readonly int _firstRow = 1;
        private readonly int _lastRow = 5;

        private Worksheet? _currentSheet;

        public bool StartSheet(Worksheet sheet)
        {
            _currentSheet = sheet;
            return true; // process this sheet
        }

        public bool StartRow(int rowIndex)
        {
            // Process only rows within the defined range
            return rowIndex >= _firstRow && rowIndex <= _lastRow;
        }

        public bool ProcessRow(Row row)
        {
            if (_currentSheet == null) return false;

            // Determine the last column that contains data in this row by scanning cells
            int lastCol = -1;
            int maxColumns = _currentSheet.Cells.MaxColumn; // total columns in the sheet
            for (int c = 0; c < maxColumns; c++)
            {
                if (_currentSheet.Cells[row.Index, c].Value != null)
                {
                    lastCol = c;
                }
            }

            int colCount = lastCol + 1; // convert to count (0‑based index to count)

            // Update the overall maximum column count
            if (colCount > MaxColumnCount) MaxColumnCount = colCount;

            // Extract cell values for the row
            object[] values = new object[colCount];
            for (int c = 0; c < colCount; c++)
            {
                values[c] = _currentSheet.Cells[row.Index, c].Value;
            }

            SelectedRows.Add(values);
            return true;
        }

        // The following methods are required by the interface but are not used for this demo
        public bool StartCell(int columnIndex) => true;
        public bool ProcessCell(Cell cell) => true;
    }

    // Supplies the collected rows to the destination workbook during saving (LightCells mode)
    class RowsProvider : LightCellsDataProvider
    {
        private readonly List<object[]> _rows;
        private readonly int _columnCount;
        private int _currentRow = -1;
        private int _currentCell = -1;

        public RowsProvider(List<object[]> rows, int columnCount)
        {
            _rows = rows;
            _columnCount = columnCount;
        }

        public bool StartSheet(int sheetIndex)
        {
            // Only one sheet is created in the destination workbook
            return sheetIndex == 0;
        }

        public int NextRow()
        {
            _currentRow++;
            _currentCell = -1;
            return _currentRow < _rows.Count ? _currentRow : -1; // -1 signals no more rows
        }

        public void StartRow(Row row)
        {
            // No special row handling required
        }

        public int NextCell()
        {
            _currentCell++;
            return _currentCell < _columnCount ? _currentCell : -1; // -1 signals no more cells in this row
        }

        public void StartCell(Cell cell)
        {
            // Write the stored value into the cell
            cell.PutValue(_rows[_currentRow][_currentCell]);
        }

        public bool IsGatherString()
        {
            // No need to gather strings into a global pool for this example
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook (replace with an actual file path)
            string sourcePath = "source.xlsx";

            try
            {
                // Ensure the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // -----------------------------------------------------------------
                // Step 1: Load the source workbook using LightCellsDataHandler to
                //         collect only the rows we are interested in.
                // -----------------------------------------------------------------
                RowCollector collector = new RowCollector();
                LoadOptions loadOptions = new LoadOptions
                {
                    LightCellsDataHandler = collector
                };
                // Loading triggers the collector; the workbook object itself is not needed afterwards
                Workbook _ = new Workbook(sourcePath, loadOptions);

                // -----------------------------------------------------------------
                // Step 2: Prepare a LightCellsDataProvider that will feed the
                //         collected rows to a new workbook during saving.
                // -----------------------------------------------------------------
                RowsProvider provider = new RowsProvider(collector.SelectedRows, collector.MaxColumnCount);

                // -----------------------------------------------------------------
                // Step 3: Create an empty workbook (it will contain a single sheet)
                //         and save it using OoxmlSaveOptions with the provider.
                // -----------------------------------------------------------------
                Workbook destWorkbook = new Workbook(); // contains one empty worksheet by default
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
                {
                    LightCellsDataProvider = provider
                };
                string outputPath = "selected_rows_output.xlsx";
                destWorkbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Selected rows have been copied to '{outputPath}' using LightCells mode.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
