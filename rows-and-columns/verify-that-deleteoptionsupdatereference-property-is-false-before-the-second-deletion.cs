// Title: C# example: verify DeleteOptions.UpdateReference is false before deleting a column with Aspose.Cells
// Common Searches: Aspose.Cells C# verify DeleteOptions.UpdateReference before deleting a column | how to prevent formula references from shifting when deleting columns with Aspose.Cells | C# example of DeleteOptions.UpdateReference set to false for column deletion | delete column without updating formulas using Aspose.Cells DeleteColumns | check DeleteOptions.UpdateReference property value prior to DeleteColumns in .NET
// Tags: DeleteOptions.UpdateReference false | DeleteColumns without reference changes | Aspose.Cells preserve formula references | C# workbook column removal example | DeleteRows then DeleteColumns Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a workbook, fills cells A1:B3, deletes the second row with default DeleteOptions (which updates references), then creates a DeleteOptions instance with UpdateReference set to false, checks the flag, and deletes the second column without updating any formula references before saving the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["A2"].PutValue(2);
        sheet.Cells["A3"].PutValue(3);
        sheet.Cells["B1"].PutValue(10);
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["B3"].PutValue(30);

        // First deletion: delete the second row (index 1) using default DeleteOptions
        DeleteOptions firstDeleteOptions = new DeleteOptions(); // UpdateReference defaults to true
        sheet.Cells.DeleteRows(1, 1, firstDeleteOptions);

        // Prepare DeleteOptions for the second deletion
        DeleteOptions secondDeleteOptions = new DeleteOptions();
        secondDeleteOptions.UpdateReference = false; // Explicitly set to false

        // Verify that UpdateReference is false before performing the second deletion
        if (!secondDeleteOptions.UpdateReference)
        {
            // Second deletion: delete the second column (original index 1) without updating references
            sheet.Cells.DeleteColumns(1, 1, secondDeleteOptions);
        }

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
