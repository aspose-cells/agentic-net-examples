// Title: Preserve Cell Comments When Saving with LightCellsDataProvider (Aspose.Cells .NET)
// Description: Demonstrates how to create a custom LightCellsDataProvider that streams the used range of the first worksheet, adds cell comments, configures OoxmlSaveOptions, and saves the workbook to XLSX while keeping all comments intact.
// Keywords: Aspose.Cells LightCellsDataProvider | save workbook with comments .NET | preserve cell annotations | stream large worksheet Aspose | OoxmlSaveOptions comments | C# Aspose.Cells example | export XLSX with annotations | lightweight cell streaming
// Common Searches: how to keep cell comments when using LightCellsDataProvider Aspose.Cells | custom LightCellsDataProvider example C# preserving comments | save workbook with comments using OoxmlSaveOptions | Aspose.Cells streaming export keep annotations | GitHub Aspose.Cells LightCellsDataProvider sample
// Developer Intent: The developer needs to save a workbook via a custom LightCellsDataProvider while ensuring that any cell comments are retained in the resulting XLSX file.
// Use Cases: Export massive worksheets to XLSX without loading the full workbook into memory, yet retain user comments for review. | Create a lightweight copy of a sheet for reporting that includes both data values and their annotations. | Implement a streaming pipeline that archives spreadsheet data together with comments for compliance purposes.
// AI Prompts: Show how to extend CustomLightCellsDataProvider to copy comment objects along with cell values. | Provide a step‑by‑step guide for saving a workbook with comments using LightCellsDataProvider and OoxmlSaveOptions in Aspose.Cells for .NET. | Explain methods to verify that comments are preserved after saving a workbook with a LightCellsDataProvider.

using System;
using Aspose.Cells;

namespace AsposeCellsLightCellsCommentsDemo
{
    // Custom LightCellsDataProvider that streams all cells of the first worksheet.
    // Demonstrates how to create a custom LightCellsDataProvider that streams the used range of the first worksheet, adds cell comments, configures OoxmlSaveOptions, and saves the workbook to XLSX while keeping all comments intact.
    public class CustomLightCellsDataProvider : LightCellsDataProvider
    {
        private readonly Worksheet _sourceSheet;
        private readonly int _maxRow;
        private readonly int _maxColumn;
        private int _currentRow = -1;
        private int _currentColumn = -1;

        public CustomLightCellsDataProvider(Worksheet sourceSheet)
        {
            _sourceSheet = sourceSheet;
            // Determine the range of used cells.
            _maxRow = sourceSheet.Cells.MaxDataRow;
            _maxColumn = sourceSheet.Cells.MaxDataColumn;
        }

        // Process only the first sheet.
        public bool StartSheet(int sheetIndex)
        {
            return sheetIndex == 0;
        }

        // Return the next row index to be saved, or -1 when done.
        public int NextRow()
        {
            if (_currentRow < _maxRow)
            {
                _currentRow++;
                _currentColumn = -1; // reset column for new row
                return _currentRow;
            }
            return -1;
        }

        // No special row handling needed.
        public void StartRow(Row row) { }

        // Return the next column index within the current row, or -1 when done.
        public int NextCell()
        {
            if (_currentColumn < _maxColumn)
            {
                _currentColumn++;
                return _currentColumn;
            }
            return -1;
        }

        // Provide the cell value from the original worksheet.
        public void StartCell(Cell cell)
        {
            // Copy the original cell value to the cell being saved.
            cell.PutValue(_sourceSheet.Cells[_currentRow, _currentColumn].Value);
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
            // Create a new workbook and access the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data.
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.25);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.75);

            // Add comments to a few cells.
            int commentIdx1 = sheet.Comments.Add("A2");
            Comment comment1 = sheet.Comments[commentIdx1];
            comment1.Note = "Fresh apples";
            comment1.Author = "Alice";

            int commentIdx2 = sheet.Comments.Add("B3");
            Comment comment2 = sheet.Comments[commentIdx2];
            comment2.Note = "Discounted price";
            comment2.Author = "Bob";

            // Prepare OoxmlSaveOptions with the custom LightCellsDataProvider.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new CustomLightCellsDataProvider(sheet)
            };

            // Save the workbook; comments are preserved in the output file.
            workbook.Save("WorkbookWithComments.xlsx", saveOptions);
        }
    }
}
