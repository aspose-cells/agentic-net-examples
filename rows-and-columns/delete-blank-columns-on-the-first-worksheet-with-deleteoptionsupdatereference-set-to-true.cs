// Title: C# – Delete Blank Columns on First Worksheet with UpdateReference using Aspose.Cells
// Description: Demonstrates how to create a workbook, insert data leaving a column empty, configure DeleteOptions.UpdateReference = true, call Cells.DeleteBlankColumns on the first worksheet, and save the file. The operation removes all empty columns while automatically adjusting formulas and references.
// Keywords: Aspose.Cells delete blank columns | DeleteBlankColumns UpdateReference | C# Aspose.Cells remove empty columns | Aspose.Cells .NET delete columns | update cell references after column deletion
// Common Searches: Aspose.Cells delete empty columns C# example | How to keep formulas updated when deleting columns in Aspose.Cells | DeleteBlankColumns with UpdateReference option | Remove blank columns from first worksheet Aspose.Cells
// Developer Intent: Remove every empty column from the first worksheet and automatically update all cell references and formulas.
// Use Cases: Clean up generated reports by eliminating columns that contain no data before exporting. | Maintain formula integrity when users leave optional columns blank in a template. | Pre‑process imported spreadsheets to strip out blank columns while preserving chart data ranges.
// AI Prompts: Write C# code using Aspose.Cells to delete all blank columns on the first worksheet and update formulas. | Show how to set DeleteOptions.UpdateReference = true with Cells.DeleteBlankColumns. | Provide an Aspose.Cells example that removes empty columns and saves the workbook as XLSX.

using System;
using Aspose.Cells;

namespace AsposeCellsDeleteBlankColumnsDemo
{
    // Demonstrates how to create a workbook, insert data leaving a column empty, configure DeleteOptions.UpdateReference = true, call Cells.DeleteBlankColumns on the first worksheet, and save the file. The operation removes all empty columns while automatically adjusting formulas and references.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add sample data with a blank column (column B will be blank)
            cells["A1"].PutValue("Column A");
            cells["C1"].PutValue("Column C"); // Column B is intentionally left blank
            cells["A2"].PutValue(1);
            cells["C2"].PutValue(3);

            // Set up delete options to update references after deletion
            DeleteOptions options = new DeleteOptions
            {
                UpdateReference = true
            };

            // Delete all blank columns on the worksheet using the specified options
            cells.DeleteBlankColumns(options);

            // Save the modified workbook
            workbook.Save("DeletedBlankColumns.xlsx", SaveFormat.Xlsx);
        }
    }
}
