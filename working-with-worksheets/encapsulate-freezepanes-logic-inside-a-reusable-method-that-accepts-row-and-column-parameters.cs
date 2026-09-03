// Title: Create a reusable C# method to freeze panes at a specific row and column using Aspose.Cells
// AI Prompts: Write a C# helper that accepts a Worksheet, row index, and column index and calls Worksheet.FreezePanes to lock the designated rows and columns. | Refactor existing Aspose.Cells code by extracting the freeze‑pane call into a parameterized function that can be reused across multiple worksheets. | Generate sample code that demonstrates invoking the reusable FreezePane method and then saving the workbook as an Excel file.
// Common Searches: how to use Aspose.Cells FreezePanes with custom row and column in C# | C# method to encapsulate worksheet freeze panes in Aspose.Cells | parameterized freeze panes example Aspose.Cells .NET | freeze rows and columns programmatically using Aspose.Cells C# | reuse freeze pane logic across multiple worksheets Aspose.Cells
// Tags: Aspose.Cells FreezePanes method | custom worksheet freeze panes utility | zero‑based indexing Aspose.Cells | freeze panes Excel generation C# | reusable worksheet freeze logic

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, and calls a custom FreezePane method with row 5 and column 3. The FreezePane method encapsulates the Aspose.Cells FreezePanes call, using zero‑based indexes, and the workbook is saved as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze panes at row 5, column 3 using the reusable method
            FreezePane(sheet, 5, 3);

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Reusable method that freezes panes at the specified row and column
    static void FreezePane(Worksheet sheet, int row, int column)
    {
        // Aspose.Cells uses zero‑based indexes for rows and columns.
        // The FreezePanes method requires the split row/column and the number of rows/columns to freeze.
        // Here we freeze rows above 'row' and columns left of 'column'.
        sheet.FreezePanes(row, column, row, column);
    }
}
