// Title: Freeze specific rows and columns in an Excel worksheet using Aspose.Cells for .NET with command‑line arguments
// AI Prompts: Generate C# code that reads row and column numbers from command‑line inputs and applies Worksheet.FreezePanes with Aspose.Cells. | Add validation for the command‑line values and default to no freeze when the inputs are missing or invalid. | Show how to modify the FreezePanes call to freeze a fixed number of rows and columns that differ from the split position.
// Common Searches: asp.net console app freeze panes using Aspose.Cells command line arguments | c# how to set freeze row and column in Excel file with Aspose.Cells | using Worksheet.FreezePanes with user‑provided indices in a .NET application | example of freezing top rows and left columns in generated Excel workbook via Aspose.Cells | command line driven freeze pane configuration in Aspose.Cells C#
// Tags: Aspose.Cells FreezePanes API C# | command‑line row column freeze Aspose.Cells | programmatic Excel freeze pane .NET | set freeze panes from args Aspose.Cells | freeze rows and columns in generated workbook

using System;
using Aspose.Cells;

// The console program reads optional row and column indices from the command line, creates a new workbook, fills sample data, applies Worksheet.FreezePanes with the supplied indices on the first worksheet, and saves the result as FrozenPaneOutput.xlsx.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Parse command‑line arguments for row and column freeze values.
            // If arguments are missing or invalid, defaults to 0 (no freeze).
            int freezeRow = 0;
            int freezeColumn = 0;

            if (args.Length >= 1 && int.TryParse(args[0], out int parsedRow))
                freezeRow = parsedRow;

            if (args.Length >= 2 && int.TryParse(args[1], out int parsedColumn))
                freezeColumn = parsedColumn;

            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (optional, just to visualize the freeze effect).
            for (int i = 0; i < 20; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Apply FreezePanes.
            // FreezePanes(row, column, totalRows, totalColumns) freezes 'totalRows' rows above 'row'
            // and 'totalColumns' columns left of 'column'.
            sheet.FreezePanes(freezeRow, freezeColumn, freezeRow, freezeColumn);

            // Save the workbook.
            workbook.Save("FrozenPaneOutput.xlsx");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
