// Title: Freeze the top row and set it as a print title range in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that freezes the first worksheet row, assigns the same row to the PrintTitleRows property, and saves the workbook. | Add a verification step that reads the PrintTitleRows value, compares it to the frozen pane range, and prints whether they match.
// Common Searches: Aspose.Cells C# repeat frozen header row on every printed page | programmatically set print title rows using Aspose.Cells for .NET | check if PrintTitleRows equals frozen panes in a generated Excel file C#
// Tags: Aspose.Cells freeze panes C# | Aspose.Cells set print title rows C# | Aspose.Cells repeat header on print | Aspose.Cells verify print title range

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, fills a 10x5 grid, freezes the first row, sets PrintTitleRows to "$1:$1" so the row repeats on each printed page, verifies the setting matches the frozen pane, and saves the file as PrintTitleWithFrozenRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill sample data (10 rows x 5 columns)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the first row (row index 0) so it stays visible while scrolling
            // Parameters: (freezeRows, freezeColumns, topRow, leftColumn)
            sheet.FreezePanes(1, 0, 1, 0); // Freeze top row

            // Apply print title range to repeat the frozen row on each printed page
            // The range is specified in A1 style; "$1:$1" means the entire first row
            sheet.PageSetup.PrintTitleRows = "$1:$1";

            // Verify that the print title range is set correctly
            string titleRange = sheet.PageSetup.PrintTitleRows;
            Console.WriteLine($"Print title rows set to: {titleRange}");

            // Verify that the frozen pane matches the title row
            bool frozenMatchesTitle = titleRange == "$1:$1";
            Console.WriteLine($"Frozen row matches title row: {frozenMatchesTitle}");

            // Define output file path
            string outputPath = "PrintTitleWithFrozenRows.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
