// Title: Aspose.Cells .NET: Delete a Column While Preserving Original Formula References (UpdateReference = false)
// Description: Demonstrates how to configure DeleteOptions.UpdateReference = false in Aspose.Cells for .NET, delete column A, and keep existing formulas unchanged before saving the workbook.
// Keywords: Aspose.Cells DeleteOptions | UpdateReference false | preserve formulas .NET | delete column without updating references | C# Aspose.Cells example
// Common Searches: Aspose.Cells delete column keep formulas | DeleteOptions.UpdateReference false sample | how to prevent formula update when deleting columns Aspose | C# Aspose.Cells preserve formula references
// Developer Intent: Remove a column from a worksheet without altering the cell formulas that reference it.
// Use Cases: Cleaning up data layouts while maintaining downstream calculations. | Removing helper columns before exporting a workbook to external systems. | Batch deleting multiple columns without breaking dependent formulas.
// AI Prompts: Write C# code using Aspose.Cells to delete rows and keep all formula references intact with DeleteOptions.UpdateReference = false. | Show an Aspose.Cells .NET example that deletes a range of columns while preserving original formulas. | Explain the impact of DeleteOptions.UpdateReference on formula behavior when columns are removed in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDeleteOptionsDemo
{
    // Demonstrates how to configure DeleteOptions.UpdateReference = false in Aspose.Cells for .NET, delete column A, and keep existing formulas unchanged before saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data and formulas
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);
            cells["C1"].Formula = "=A1+B1"; // Formula that references column A

            cells["A2"].PutValue(30);
            cells["B2"].PutValue(40);
            cells["C2"].Formula = "=A2+B2";

            // Create DeleteOptions and set UpdateReference to false
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = false // Preserve original formulas after deletion
            };

            // Delete column A (index 0) using the options
            sheet.Cells.DeleteColumns(0, 1, deleteOptions);

            // At this point, formulas in C1 and C2 still refer to the original cells (now shifted)
            Console.WriteLine("Formula in C1 after deletion: " + cells["C1"].Formula);
            Console.WriteLine("Formula in C2 after deletion: " + cells["C2"].Formula);

            // Save the workbook
            workbook.Save("DeleteOptions_UpdateReference_False.xlsx");
        }
    }
}
