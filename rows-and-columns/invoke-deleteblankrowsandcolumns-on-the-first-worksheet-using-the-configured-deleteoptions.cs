// Title: C# – Delete Blank Rows & Columns in the First Worksheet with Aspose.Cells DeleteOptions
// Description: Creates a workbook, adds sample data with intentional empty rows and columns, configures DeleteOptions (UpdateReference = true), then calls DeleteBlankRows and DeleteBlankColumns on the first worksheet's Cells collection and saves the result.
// Keywords: Aspose.Cells | C# | .NET | DeleteBlankRows | DeleteBlankColumns | DeleteOptions | UpdateReference | remove empty rows | remove empty columns | workbook cleanup
// Common Searches: Aspose.Cells delete blank rows C# | How to remove empty columns using DeleteOptions | DeleteBlankRows DeleteBlankColumns example .NET | UpdateReference option Aspose.Cells | Clean up first worksheet Aspose.Cells
// Developer Intent: Remove all empty rows and columns from the first worksheet while keeping cell references and formulas accurate.
// Use Cases: Trim generated reports by deleting blank rows/columns and automatically updating formulas. | Prepare data for export by eliminating unnecessary empty space from the first sheet. | Apply a consistent DeleteOptions configuration across multiple worksheets to standardize cleanup.
// AI Prompts: Generate C# code that uses Aspose.Cells to delete all blank rows and columns in the first worksheet with UpdateReference enabled. | Show how to set DeleteOptions.UpdateReference = true and apply it to DeleteBlankRows and DeleteBlankColumns methods. | Explain the difference between DeleteBlankRows, DeleteBlankColumns, and a combined blank‑space removal in Aspose.Cells .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsDeleteBlankRowsAndColumns
{
    // Creates a workbook, adds sample data with intentional empty rows and columns, configures DeleteOptions (UpdateReference = true), then calls DeleteBlankRows and DeleteBlankColumns on the first worksheet's Cells collection and saves the result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data with blank rows and columns
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Data 1");
            // Row 3 is blank
            cells["A4"].PutValue("Data 2"); // Blank row at A3
            // Column B is blank
            cells["C1"].PutValue("Extra");

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
