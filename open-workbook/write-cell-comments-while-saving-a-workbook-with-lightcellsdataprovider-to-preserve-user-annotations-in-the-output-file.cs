// Title: Preserve Cell Comments When Saving with LightCellsDataProvider – Aspose.Cells C# Example
// Description: Demonstrates how to create a workbook, add comments, implement a custom LightCellsDataProvider that streams cell values, configure OoxmlSaveOptions, and save the file as XLSX while keeping all cell comments intact.
// Keywords: Aspose.Cells | LightCellsDataProvider | C# | preserve comments | save workbook with comments | streaming export | low memory Excel export | OoxmlSaveOptions | custom data provider
// Common Searches: How to keep cell comments using LightCellsDataProvider in Aspose.Cells | Aspose.Cells example for saving workbook with comments | Custom LightCellsDataProvider C# sample | Export large Excel sheet with comments preserved | LightCellsDataProvider comment retention
// Developer Intent: Save an Excel workbook with a custom LightCellsDataProvider while ensuring that all cell comments are retained in the output file.
// Use Cases: Export massive worksheets to XLSX with minimal memory usage and retain reviewer notes. | Build a streaming report generator that copies data from a template and keeps annotation comments for audit purposes. | Create a low‑memory service that writes data to a new workbook while preserving user comments for downstream processing.
// AI Prompts: Show a C# code snippet that uses Aspose.Cells LightCellsDataProvider to copy a worksheet and keep all comments. | Explain how to extend CustomLightCellsDataProvider to also copy comment objects during streaming. | Provide steps to configure OoxmlSaveOptions with LightCellsDataProvider and verify that comments appear in the saved XLSX.

using System;
using Aspose.Cells;

namespace LightCellsCommentsDemo
{
    // Custom LightCellsDataProvider that streams cell values from an existing worksheet.
    // Demonstrates how to create a workbook, add comments, implement a custom LightCellsDataProvider that streams cell values, configure OoxmlSaveOptions, and save the file as XLSX while keeping all cell comments intact.
    public class CustomLightCellsDataProvider : LightCellsDataProvider
    {
        private readonly Worksheet _sourceWorksheet;
        private int _currentRow = -1;
        private int _currentColumn = -1;
        private readonly int _maxRow;
        private readonly int _maxColumn;

        public CustomLightCellsDataProvider(Worksheet sourceWorksheet)
        {
            _sourceWorksheet = sourceWorksheet;
            // Determine the range of used cells to stream.
            _maxRow = sourceWorksheet.Cells.MaxDataRow;
            _maxColumn = sourceWorksheet.Cells.MaxDataColumn;
        }

        // Process only the first sheet.
        public bool StartSheet(int sheetIndex)
        {
            return sheetIndex == 0;
        }

        // Return the next row index to be saved, or -1 when finished.
        public int NextRow()
        {
            _currentRow++;
            _currentColumn = -1;
            return _currentRow <= _maxRow ? _currentRow : -1;
        }

        // No special row initialization required.
        public void StartRow(Row row) { }

        // Return the next column index within the current row, or -1 when finished.
        public int NextCell()
        {
            _currentColumn++;
            return _currentColumn <= _maxColumn ? _currentColumn : -1;
        }

        // Provide the cell value from the source worksheet.
        public void StartCell(Cell cell)
        {
            Cell srcCell = _sourceWorksheet.Cells[_currentRow, _currentColumn];
            if (srcCell != null && srcCell.Type != CellValueType.IsNull)
            {
                cell.PutValue(srcCell.Value);
            }
        }

        // Indicate that string values should be gathered into the global string pool.
        public bool IsGatherString()
        {
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data.
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1.25);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(0.80);

            // Add comments to a few cells.
            int commentIdx1 = worksheet.Comments.Add("A2");
            Comment comment1 = worksheet.Comments[commentIdx1];
            comment1.Note = "Fresh apples";
            comment1.Author = "Alice";

            int commentIdx2 = worksheet.Comments.Add("B3");
            Comment comment2 = worksheet.Comments[commentIdx2];
            comment2.Note = "Discounted price";
            comment2.Author = "Bob";

            // Configure OoxmlSaveOptions to use the custom LightCellsDataProvider.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new CustomLightCellsDataProvider(worksheet)
            };

            // Save the workbook; comments are preserved in the output file.
            workbook.Save("WorkbookWithComments.xlsx", saveOptions);

            Console.WriteLine("Workbook saved with comments using LightCellsDataProvider.");
        }
    }
}
