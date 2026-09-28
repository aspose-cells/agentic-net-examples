// Title: Delete a column in Aspose.Cells for .NET while preserving original formulas by setting DeleteOptions.UpdateReference to false
// AI Prompts: Write C# code that deletes a specific column in an Aspose.Cells workbook and keeps all existing formulas unchanged by configuring DeleteOptions.UpdateReference = false. | Show how to use Aspose.Cells DeleteOptions to remove rows or columns without adjusting any formula references in a .NET application.
// Common Searches: Aspose.Cells C# delete column without updating formulas | How to keep formula references intact after deleting a column in Aspose.Cells | DeleteOptions.UpdateReference false example for preserving formulas | Prevent formula recalculation when removing rows in Aspose.Cells .NET | Preserve original cell references after column deletion using Aspose.Cells API
// Tags: delete column without formula update Aspose.Cells | preserve formula references Aspose.Cells | disable formula reference adjustment Aspose.Cells | C# DeleteOptions usage Aspose.Cells | column removal keep original formulas .NET

using System;
using Aspose.Cells;

// Creates a workbook, adds values and a formula, deletes column A using DeleteOptions with UpdateReference set to false so the formula in C1 remains unchanged, and saves the file as PreserveFormulas.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some cells with values and a formula that references column A
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["B1"].PutValue(20);
        worksheet.Cells["C1"].Formula = "=A1+B1";

        // Create DeleteOptions and set UpdateReference to false
        // This ensures that formulas referencing deleted cells are NOT updated
        DeleteOptions deleteOptions = new DeleteOptions
        {
            UpdateReference = false
        };

        // Delete column A (index 0) using the DeleteOptions
        // The formula in C1 will remain "=A1+B1" even though column A is removed
        worksheet.Cells.DeleteColumns(0, 1, deleteOptions);

        // Save the workbook to a file
        workbook.Save("PreserveFormulas.xlsx");
    }
}
