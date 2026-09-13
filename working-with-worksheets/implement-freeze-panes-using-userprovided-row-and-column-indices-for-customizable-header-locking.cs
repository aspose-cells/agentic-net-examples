// Title: Freeze custom rows and columns in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that freezes a user‑specified number of rows and columns in an Aspose.Cells workbook and saves the file. | Show how to use the Worksheet.FreezePanes method with dynamic row and column indices and include basic exception handling. | Create a snippet that reads row and column values from variables, applies FreezePanes, and writes the result to an .xlsx file.
// Common Searches: how to freeze the first two rows and first column in Excel using Aspose.Cells C# | Aspose.Cells programmatically lock header rows based on user input | C# example for setting custom freeze panes in an .xlsx workbook with Aspose | using Worksheet.FreezePanes to define scrollable area in .NET
// Tags: Aspose.Cells worksheet FreezePanes C# | custom row column freeze Excel .NET | programmatic header locking Aspose.Cells | freeze panes with user-defined indices | save workbook after applying FreezePanes

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, and calls Worksheet.FreezePanes with user‑provided row and column counts to lock the top rows and first column. The workbook is then saved as FreezePanesResult.xlsx, with basic exception handling around the operation.
class FreezePanesExample
{
    static void Main()
    {
        try
        {
            // User‑provided indices (0‑based for the API)
            // Example: freeze first 2 rows and first column
            int freezeRow = 2;      // Number of rows to freeze
            int freezeColumn = 1;   // Number of columns to freeze

            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Apply freeze panes:
            // freezeRow and freezeColumn specify the top‑left cell of the scrollable area,
            // the same values are used for totalRows and totalColumns to define the frozen area.
            sheet.FreezePanes(freezeRow, freezeColumn, freezeRow, freezeColumn);

            // Save the workbook (lifecycle rule)
            workbook.Save("FreezePanesResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
