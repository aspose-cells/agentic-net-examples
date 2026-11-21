// Title: Delete a row in an Excel file without updating formula references using Aspose.Cells DeleteOptions in C#
// Common Searches: C# Aspose.Cells delete row keep formula references unchanged | How to set DeleteOptions.UpdateReference false when deleting rows in Aspose.Cells | Preserve original formulas after removing rows with Aspose.Cells .NET | Aspose.Cells DeleteRows example with UpdateReference set to false | Prevent formula reference shift when deleting rows using Aspose.Cells in C#
// Tags: DeleteRows DeleteOptions UpdateReference false | Aspose.Cells formula reference preservation | C# Excel row deletion without formula update | Aspose.Cells DeleteOptions configuration | Excel workbook row removal maintaining formulas

using Aspose.Cells;

// The program loads an Excel workbook, creates a DeleteOptions object with UpdateReference set to false to keep existing formulas unchanged, deletes the fifth row, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Create DeleteOptions and set UpdateReference to false
        DeleteOptions deleteOptions = new DeleteOptions();
        deleteOptions.UpdateReference = false; // Preserve original formulas

        // Example: delete the 5th row (zero‑based index 4) using the options
        sheet.Cells.DeleteRows(4, 1, deleteOptions);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
