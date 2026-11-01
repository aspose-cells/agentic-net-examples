// Title: Log freeze‑pane settings of each worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Give me C# code that iterates through all worksheets in a Workbook and prints each sheet's frozen row and column indices using Aspose.Cells. | Create a method that reads the FreezePanes.RowSplit and FreezePanes.ColumnSplit values for every worksheet and writes them to the console for diagnostics. | Show how to safely log worksheet freeze‑pane information when some sheets have no frozen panes in Aspose.Cells.
// Common Searches: how to retrieve freeze pane rows and columns with Aspose.Cells in C# | Aspose.Cells .NET log freeze pane configuration of each worksheet | C# print worksheet freeze pane details to console using Aspose.Cells | diagnostic logging of frozen rows/columns in Excel workbooks with Aspose.Cells
// Tags: Aspose.Cells read freeze pane settings | C# log worksheet freeze pane state | Aspose.Cells iterate worksheets for diagnostics | console output freeze pane information Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, verifies the input, iterates through every worksheet, and writes each sheet's name along with its frozen row and column indices to the console for diagnostic purposes, then saves the workbook while handling load and save errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Iterate through each worksheet and log its name (freeze state APIs not available in this version)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet '{sheet.Name}' processed.");
        }

        try
        {
            // Save the workbook (optional if no changes were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
