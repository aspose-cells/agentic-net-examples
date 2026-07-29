// Title: C# – Delete Blank Rows & Columns on First Worksheet with DeleteOptions (Aspose.Cells)
// Description: Creates an in‑memory workbook, adds sample data with intentional empty rows and columns, configures DeleteOptions (UpdateReference = true), invokes Cells.DeleteBlankRows and Cells.DeleteBlankColumns on the first worksheet, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# | DeleteBlankRows | DeleteBlankColumns | DeleteOptions | UpdateReference | remove empty rows | remove empty columns | blank row deletion | blank column deletion | Excel cleanup .NET
// Common Searches: Aspose.Cells delete blank rows C# | Aspose.Cells delete blank columns with DeleteOptions | How to remove empty rows and columns in .NET Excel | Update formulas after deleting rows Aspose.Cells | DeleteBlankRowsAndColumns example
// Developer Intent: Remove all empty rows and columns from the first worksheet while preserving formula references using DeleteOptions.
// Use Cases: Trim generated workbooks to eliminate gaps and reduce file size before distribution. | Prepare data ranges for charts or reports by eliminating blank rows/columns that could skew calculations. | Maintain accurate formula references after cleanup by enabling UpdateReference in DeleteOptions.
// AI Prompts: Show C# code that deletes blank rows and columns on the first worksheet using DeleteOptions with UpdateReference in Aspose.Cells. | Provide an Aspose.Cells example that removes empty rows/columns, updates cell references, and saves the workbook as XLSX. | Explain how DeleteOptions.UpdateReference affects formulas when blank rows or columns are removed in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates an in‑memory workbook, adds sample data with intentional empty rows and columns, configures DeleteOptions (UpdateReference = true), invokes Cells.DeleteBlankRows and Cells.DeleteBlankColumns on the first worksheet, and saves the result as an XLSX file.
    public class DeleteBlankRowsAndColumnsDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data with blank rows and columns
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Data1");
            // Row 3 is intentionally left blank
            cells["A4"].PutValue("Data2"); // Blank row at A3
            // Column B is intentionally left blank
            cells["C1"].PutValue("Extra");
            cells["C2"].PutValue(123);

            // Configure DeleteOptions (e.g., update references after deletion)
            DeleteOptions options = new DeleteOptions
            {
                UpdateReference = true
            };

            // Delete all blank rows using the configured options
            cells.DeleteBlankRows(options);

            // Delete all blank columns using the same options
            cells.DeleteBlankColumns(options);

            // Save the modified workbook
            workbook.Save("DeletedBlankRowsAndColumns.xlsx", SaveFormat.Xlsx);
        }
    }
}
