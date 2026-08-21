// Title: Automatically adjust formulas when deleting columns with DeleteOptions.UpdateReference in Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds values and formulas referencing column A, sets DeleteOptions.UpdateReference = true, deletes column A, and saves the file, showing how formulas are recalculated to reflect the new column positions.
// Keywords: Aspose.Cells DeleteOptions.UpdateReference | C# delete column keep formulas | Aspose.Cells formula adjustment | DeleteColumns with reference update | .NET spreadsheet API delete column | automatic formula shift Aspose
// Common Searches: Aspose.Cells DeleteOptions.UpdateReference true example | how to keep formulas after deleting a column in Aspose.Cells | DeleteColumns with formula update .NET | update cell references when removing rows Aspose.Cells
// Developer Intent: Enable the UpdateReference flag so that any formulas referencing deleted rows or columns are automatically rewritten to maintain correct calculations.
// Use Cases: Delete a column that contains raw data while preserving dependent calculations in other sheets. | Remove a block of rows and have all related formulas shift without manual editing. | Clean up a worksheet programmatically by eliminating unused columns without breaking existing formulas.
// AI Prompts: Provide a C# snippet that deletes multiple columns with DeleteOptions.UpdateReference set to true in Aspose.Cells. | Show how to delete a row range and automatically update formulas using DeleteOptions.UpdateReference. | Compare formula behavior in Aspose.Cells when DeleteOptions.UpdateReference is false versus true after column deletion.

using System;
using Aspose.Cells;

namespace AsposeCellsDeleteOptionsDemo
{
    // C# example that creates a workbook, adds values and formulas referencing column A, sets DeleteOptions.UpdateReference = true, deletes column A, and saves the file, showing how formulas are recalculated to reflect the new column positions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data and formulas that reference column A
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);
            cells["C1"].Formula = "=A1+B1"; // Formula references column A

            cells["A2"].PutValue(30);
            cells["B2"].PutValue(40);
            cells["C2"].Formula = "=A2+B2";

            // Create DeleteOptions and enable reference updating
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = true // Enable formula reference updates
            };

            // Delete column A (index 0) using the options
            // Formulas in column C will be adjusted automatically
            sheet.Cells.DeleteColumns(0, 1, deleteOptions);

            // Save the workbook to verify that formulas have been updated
            workbook.Save("DeleteOptions_UpdateReference_True.xlsx");
        }
    }
}
