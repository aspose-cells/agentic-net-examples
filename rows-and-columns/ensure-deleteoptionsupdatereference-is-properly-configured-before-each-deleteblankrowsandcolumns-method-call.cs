// Title: Configure DeleteBlankOptions.UpdateReference for DeleteBlankRows and DeleteBlankColumns in Aspose.Cells .NET
// Description: C# example that creates a workbook, adds data and a formula, then sets DeleteBlankOptions.UpdateReference = true (with EmptyStringAsBlank). The configured options are used to remove blank rows and columns while automatically adjusting any formula references, and the workbook is saved as Result.xlsx.
// Keywords: Aspose.Cells DeleteBlankOptions | UpdateReference true | DeleteBlankRows C# | DeleteBlankColumns C# | preserve formula references | treat empty strings as blanks | Aspose.Cells .NET example | remove blank rows columns | GitHub Aspose.Cells sample | C# spreadsheet automation
// Common Searches: Aspose.Cells keep formulas when deleting blank rows | Set UpdateReference for DeleteBlankRows in C# | DeleteBlankColumns without breaking references Aspose.Cells | How to treat empty strings as blanks in Aspose.Cells | C# code sample for DeleteBlankOptions
// Developer Intent: Enable UpdateReference on DeleteBlankOptions before calling DeleteBlankRows or DeleteBlankColumns so that cell references are recalculated automatically.
// Use Cases: Remove empty rows while ensuring dependent formulas stay correct. | Delete unused columns and have all related calculations update instantly. | Consider cells containing "" as blank during cleanup operations.
// AI Prompts: Provide C# code that sets DeleteBlankOptions.UpdateReference to true and deletes blank rows and columns with Aspose.Cells. | Explain why omitting UpdateReference can break formulas when using DeleteBlankRows or DeleteBlankColumns. | Show an Aspose.Cells .NET example that treats empty strings as blanks while cleaning a worksheet.

using Aspose.Cells;
using System;

// C# example that creates a workbook, adds data and a formula, then sets DeleteBlankOptions.UpdateReference = true (with EmptyStringAsBlank). The configured options are used to remove blank rows and columns while automatically adjusting any formula references, and the workbook is saved as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some data, including blank rows/columns and a formula that references them
        cells["A1"].PutValue(10);
        cells["B1"].PutValue(20);
        cells["C1"].Formula = "=A1+B1"; // Formula will be updated when columns are deleted
        cells["A2"].PutValue("");      // Blank row
        cells["B2"].PutValue("");      // Blank column
        cells["C2"].PutValue("");

        // Prepare DeleteBlankOptions with UpdateReference set to true
        DeleteBlankOptions options = new DeleteBlankOptions
        {
            UpdateReference = true,      // Ensure references are updated after deletion
            EmptyStringAsBlank = true    // Treat empty strings as blanks
        };

        // Delete blank rows using the configured options
        cells.DeleteBlankRows(options);

        // Delete blank columns using the same options (UpdateReference remains true)
        cells.DeleteBlankColumns(options);

        // Save the modified workbook
        workbook.Save("Result.xlsx");
    }
}
