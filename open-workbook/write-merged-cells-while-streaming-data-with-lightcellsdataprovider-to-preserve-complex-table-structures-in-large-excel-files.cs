// Title: Write Merged Header Cells While Streaming Data with LightCellsDataProvider – Aspose.Cells C# Example
// Description: Demonstrates a custom LightCellsDataProvider that streams rows from a two‑dimensional array, merges the first three cells (A1:C1) to create a header, and saves the workbook with OoxmlSaveOptions. The merged range stays intact, enabling low‑memory generation of large Excel files in .NET.
// Keywords: Aspose.Cells | LightCellsDataProvider | merged cells | streaming Excel data | large workbook .NET | C# Excel export | low memory Excel generation | OoxmlSaveOptions | custom data provider | Excel header merge
// Common Searches: Aspose.Cells keep merged cells when using LightCellsDataProvider | stream large Excel file with merged header C# | custom LightCellsDataProvider example | write merged cells while streaming data Aspose.Cells | low memory Excel export .NET merged cells
// Developer Intent: Generate an Excel workbook that streams row‑by‑row data via a custom LightCellsDataProvider while preserving a merged header row.
// Use Cases: Export millions of database rows to Excel with a title that spans several columns, avoiding high memory consumption. | Create financial statements where the report title is merged across columns and the detail rows are written on the fly. | Build large CSV‑to‑XLSX converters that need to retain template formatting such as merged header cells.
// AI Prompts: Show a C# code sample that uses Aspose.Cells LightCellsDataProvider to write data into a worksheet with a merged header across A1:C1. | Explain how to add custom styling (font, background) to the merged header cells while streaming data with LightCellsDataProvider. | Provide steps to configure OoxmlSaveOptions for low‑memory saving of a workbook that contains predefined merged ranges.

using System;
using Aspose.Cells;

namespace LightCellsMergedCellsDemo
{
    // Custom LightCellsDataProvider that streams data row by row and cell by cell.
    // It supplies values from a simple two‑dimensional array.
    // Demonstrates a custom LightCellsDataProvider that streams rows from a two‑dimensional array, merges the first three cells (A1:C1) to create a header, and saves the workbook with OoxmlSaveOptions. The merged range stays intact, enabling low‑memory generation of large Excel files in .NET.
    public class MergedCellsDataProvider : LightCellsDataProvider
    {
        // Sample data to be written to the worksheet.
        // First row is a header that will be merged across the first three columns.
        private readonly string[,] _data = new string[,]
        {
            { "Header", "", "" },          // Header row (merged cells will cover A1:C1)
            { "ID", "Name", "Amount" },    // Column titles
            { "1", "Item A", "100" },
            { "2", "Item B", "200" },
            { "3", "Item C", "300" }
        };

        private int _currentRow = -1;
        private int _currentColumn = -1;

        // Called for each worksheet. Return true for the first sheet (index 0) to use this provider.
        public bool StartSheet(int sheetIndex)
        {
            // Process only the first worksheet.
            return sheetIndex == 0;
        }

        // Returns the next row index to be saved, or -1 when no more rows.
        public int NextRow()
        {
            _currentRow++;
            _currentColumn = -1; // Reset column index for the new row.

            // If the row index is within the data array, return it; otherwise stop.
            return _currentRow < _data.GetLength(0) ? _currentRow : -1;
        }

        // Provides the Row object for the current row (no special handling needed here).
        public void StartRow(Row row)
        {
            // Example: set a fixed row height for all rows.
            row.Height = 15;
        }

        // Returns the next column index for the current row, or -1 when the row is finished.
        public int NextCell()
        {
            _currentColumn++;
            return _currentColumn < _data.GetLength(1) ? _currentColumn : -1;
        }

        // Supplies the cell value for the current row/column.
        public void StartCell(Cell cell)
        {
            // Put the value from the data array into the cell.
            cell.PutValue(_data[_currentRow, _currentColumn]);
        }

        // Determines whether string values should be gathered into a global string pool.
        // Here we enable gathering for the header row to demonstrate the option.
        public bool IsGatherString()
        {
            // Gather strings only for the header row (row index 0).
            return _currentRow == 0;
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and obtain the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // 2. Define merged cells that span the header across columns A, B and C.
            // Merge range: firstRow = 0, firstColumn = 0, totalRows = 1, totalColumns = 3
            worksheet.Cells.Merge(0, 0, 1, 3);

            // 3. Create an instance of the custom LightCellsDataProvider.
            LightCellsDataProvider dataProvider = new MergedCellsDataProvider();

            // 4. Configure save options to use the LightCellsDataProvider.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = dataProvider
            };

            // 5. Save the workbook. The merged cells defined in step 2 are preserved,
            //    while the cell values are streamed from the provider.
            workbook.Save("MergedCellsStreamingDemo.xlsx", saveOptions);

            Console.WriteLine("Workbook saved successfully with merged cells and streamed data.");
        }
    }
}
