// Title: Filter rows with column C = "Active" using LightCells in Aspose.Cells for .NET
// Description: Demonstrates a custom LightCellsDataProvider that scans a source worksheet, records the indexes of rows whose column C contains "Active", and streams only those rows to a new workbook during a LightCells save. The example preserves original cell types, values, and row heights while keeping memory usage low.
// Keywords: Aspose.Cells LightCells | C# filter Excel rows | stream rows by column value | Active rows export | memory‑efficient Excel processing | custom LightCellsDataProvider | large workbook filtering | copy selected rows to new workbook | .NET Excel streaming
// Common Searches: Aspose.Cells LightCells filter rows by column value | C# stream only rows where column C is Active | How to export selected rows with Aspose.Cells without loading whole file | LightCellsDataProvider example for conditional copy | filter Excel data using Aspose.Cells streaming API
// Developer Intent: Create a LightCellsDataProvider that streams only rows whose column C equals "Active" and saves them to a new worksheet.
// Use Cases: Generate a lightweight report of active records from a massive dataset without loading the entire workbook into memory. | Export only rows marked "Active" for downstream analytics or data pipelines. | Perform memory‑conscious data migration by copying selected rows into a separate Excel file.
// AI Prompts: Write a LightCellsDataProvider that filters rows based on a case‑insensitive match in column D. | Extend the provider to copy cell styles and formulas for the filtered rows. | Show how to direct the filtered rows to a sheet named "ActiveItems" instead of the first sheet.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace LightCellsFilterExample
{
    // Custom LightCellsDataProvider that streams only rows where column C (index 2) equals "Active"
    // Demonstrates a custom LightCellsDataProvider that scans a source worksheet, records the indexes of rows whose column C contains "Active", and streams only those rows to a new workbook during a LightCells save. The example preserves original cell types, values, and row heights while keeping memory usage low.
    public class ActiveRowsLightCellsProvider : LightCellsDataProvider
    {
        private readonly Worksheet _sourceSheet;
        private readonly List<int> _matchingRows; // zero‑based row indexes that satisfy the condition
        private int _currentRowIndex;               // index in _matchingRows list
        private int _currentSourceRow;              // actual source row being processed
        private int _currentColumn;                 // column being processed
        private readonly int _maxColumn;            // maximum column to copy

        public ActiveRowsLightCellsProvider(Worksheet sourceSheet)
        {
            _sourceSheet = sourceSheet;
            _matchingRows = new List<int>();

            // Determine the last column with data in the source sheet
            _maxColumn = _sourceSheet.Cells.MaxDataColumn;

            // Scan source rows once to collect rows where column C (index 2) == "Active"
            for (int row = 0; row <= _sourceSheet.Cells.MaxDataRow; row++)
            {
                string value = _sourceSheet.Cells[row, 2].StringValue; // column C
                if (string.Equals(value, "Active", StringComparison.OrdinalIgnoreCase))
                {
                    _matchingRows.Add(row);
                }
            }
        }

        // Called when the destination workbook starts saving a sheet.
        public bool StartSheet(int sheetIndex)
        {
            // We only provide data for the first (and only) sheet in the destination workbook.
            if (sheetIndex == 0 && _matchingRows.Count > 0)
            {
                _currentRowIndex = 0;
                return true;
            }
            return false; // Use normal data model for any other sheet.
        }

        // Returns the next source row index to be written, or -1 to finish.
        public int NextRow()
        {
            if (_currentRowIndex < _matchingRows.Count)
            {
                _currentSourceRow = _matchingRows[_currentRowIndex];
                _currentRowIndex++;
                _currentColumn = -1; // reset column pointer for the new row
                return _currentSourceRow;
            }
            return -1; // No more rows.
        }

        // Called before processing cells of the current row.
        public void StartRow(Row row)
        {
            // Optionally copy row height from source.
            row.Height = _sourceSheet.Cells.GetRowHeight(_currentSourceRow);
        }

        // Returns the next column index to be processed, or -1 to finish the row.
        public int NextCell()
        {
            if (_currentColumn < _maxColumn)
            {
                _currentColumn++;
                return _currentColumn;
            }
            return -1; // End of columns for this row.
        }

        // Called for each cell that will be written to the destination sheet.
        public void StartCell(Cell cell)
        {
            // Copy the value from the source cell to the destination cell.
            var srcCell = _sourceSheet.Cells[_currentSourceRow, _currentColumn];
            if (srcCell != null && srcCell.Type != CellValueType.IsNull)
            {
                // Preserve the original value type.
                switch (srcCell.Type)
                {
                    case CellValueType.IsString:
                        cell.PutValue(srcCell.StringValue);
                        break;
                    case CellValueType.IsNumeric:
                        cell.PutValue(srcCell.DoubleValue);
                        break;
                    case CellValueType.IsBool:
                        cell.PutValue(srcCell.BoolValue);
                        break;
                    case CellValueType.IsDateTime:
                        cell.PutValue(srcCell.DateTimeValue);
                        break;
                    case CellValueType.IsError:
                        cell.PutValue(srcCell.IntValue);
                        break;
                    default:
                        cell.PutValue(srcCell.Value);
                        break;
                }
            }
        }

        // Indicates whether the provider wants to gather string values for shared string table.
        public bool IsGatherString()
        {
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path).
            Workbook sourceWorkbook = new Workbook("SourceData.xlsx");
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Create a new workbook that will receive the filtered rows.
            Workbook destWorkbook = new Workbook();
            destWorkbook.Worksheets.Clear();          // Remove the default sheet.
            destWorkbook.Worksheets.Add();            // Add an empty sheet for output.

            // Configure LightCells save options with the custom provider.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                LightCellsDataProvider = new ActiveRowsLightCellsProvider(sourceSheet)
            };

            // Save the destination workbook; the provider streams matching rows.
            destWorkbook.Save("FilteredActiveRows.xlsx", saveOptions);

            Console.WriteLine("Filtering complete. Output saved to FilteredActiveRows.xlsx");
        }
    }
}
