// Title: How to delete rows while automatically updating Excel formula references using DeleteOptions.UpdateReference in Aspose.Cells for .NET
// Common Searches: Aspose.Cells delete rows and keep formulas correct in C# | set DeleteOptions.UpdateReference to true for row deletion in .NET | C# code to remove multiple rows and adjust Excel formula references with Aspose.Cells | using DeleteOptions to preserve formula links after deleting rows in a workbook | update Excel cell references after row removal with Aspose.Cells DeleteOptions
// Tags: DeleteOptions.UpdateReference usage in Aspose.Cells | row deletion with automatic formula adjustment C# | preserve Excel formulas during row removal Aspose.Cells | adjust cell references after deleting rows .NET | Aspose.Cells workbook row removal preserving formulas

using Aspose.Cells;
using System;

// The example loads an Excel workbook, creates a DeleteOptions object with UpdateReference set to true so that formulas are automatically adjusted, deletes rows 5‑7 from the first worksheet, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Create DeleteOptions and enable formula reference updates
        DeleteOptions deleteOptions = new DeleteOptions();
        deleteOptions.UpdateReference = true; // formulas will adjust after deletion

        // Example: delete rows 5, 6, and 7 using the DeleteOptions
        sheet.Cells.DeleteRows(5, 3, deleteOptions);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
