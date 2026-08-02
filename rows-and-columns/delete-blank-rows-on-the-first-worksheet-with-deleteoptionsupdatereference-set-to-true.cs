// Title: Aspose.Cells for .NET – Delete Blank Rows with DeleteOptions.UpdateReference = true
// Description: C# example that creates a workbook, inserts data with empty rows, configures DeleteOptions to update references, removes all blank rows from the first worksheet using Cells.DeleteBlankRows, and saves the file as DeletedBlankRows.xlsx.
// Keywords: Aspose.Cells delete blank rows | DeleteOptions UpdateReference | C# remove empty rows | Cells.DeleteBlankRows method | update formulas after row deletion | .NET spreadsheet automation
// Common Searches: Aspose.Cells delete blank rows update references | C# DeleteBlankRows with DeleteOptions | how to remove empty rows in Aspose.Cells | keep formulas correct after deleting rows Aspose.Cells | DeleteOptions.UpdateReference example
// Developer Intent: Remove every empty row from the first worksheet while automatically adjusting any cell references or formulas that point to the deleted rows.
// Use Cases: Clean up generated reports by eliminating stray blank rows before exporting. | Maintain accurate formula results after row removal in financial or analytical workbooks. | Prepare data for downstream systems that cannot handle empty rows.
// AI Prompts: Write C# code using Aspose.Cells to delete all blank rows on the first worksheet with DeleteOptions.UpdateReference set to true. | Show how DeleteOptions.UpdateReference affects formulas when calling Cells.DeleteBlankRows in Aspose.Cells for .NET. | Explain step‑by‑step how to configure DeleteOptions and invoke DeleteBlankRows to keep references up‑to‑date.

using System;
using Aspose.Cells;

namespace DeleteBlankRowsExample
{
    // C# example that creates a workbook, inserts data with empty rows, configures DeleteOptions to update references, removes all blank rows from the first worksheet using Cells.DeleteBlankRows, and saves the file as DeletedBlankRows.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data with blank rows
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Data1");
            // Row 3 will be blank
            cells["A4"].PutValue("Data2");
            // Row 5 will be blank
            cells["A6"].PutValue("Data3");

            // Set up DeleteOptions with UpdateReference = true
            DeleteOptions options = new DeleteOptions
            {
                UpdateReference = true
            };

            // Delete all blank rows on the first worksheet using the options
            cells.DeleteBlankRows(options);

            // Save the workbook
            workbook.Save("DeletedBlankRows.xlsx", SaveFormat.Xlsx);
        }
    }
}
