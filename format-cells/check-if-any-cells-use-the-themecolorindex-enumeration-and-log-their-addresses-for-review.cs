// Title: Identify and log Excel cells that use ThemeColor formatting with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, scans every worksheet, and prints the addresses of cells whose font, foreground, or background style has a non‑default ThemeColor. | Create a reusable method that accepts a Worksheet object and returns a list of cell names where any ThemeColor property (Font.ThemeColor, ForegroundThemeColor, BackgroundThemeColor) is set, then output each cell together with its worksheet name. | Adapt the example to write the detected ThemeColor cell addresses to a CSV file, including columns for worksheet name, cell address, and the specific ThemeColor property that triggered the detection.
// Common Searches: C# Aspose.Cells how to find cells with theme color formatting in an Excel workbook | list all cells using ThemeColor enum in .xlsx using Aspose.Cells for .NET | detect non-default ThemeColor in cell styles with Aspose.Cells C# example | log addresses of cells that have font or fill theme colors using Aspose.Cells | iterate through worksheets and check ThemeColor properties in Aspose.Cells
// Tags: Aspose.Cells scan ThemeColor in cell styles | C# detect ThemeColor usage in Excel worksheets | log cell addresses with ThemeColor formatting | enumerate cells using ThemeColor enum Aspose.Cells | extract ThemeColor‑styled cells from .xlsx

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, iterates through each worksheet and its used cells, checks whether the cell's font, foreground, or background style uses a non‑default ThemeColor, and writes the worksheet name and cell address to the console for every match.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found – {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the specified file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Retrieve the cell's style
                    Style style = cell.GetStyle();

                    // Determine if the cell's fill or font uses a ThemeColor value
                    bool usesThemeColor =
                        style.Font.ThemeColor != default(ThemeColor) ||
                        style.ForegroundThemeColor != default(ThemeColor) ||
                        style.BackgroundThemeColor != default(ThemeColor);

                    // If a ThemeColor is used, log the cell address
                    if (usesThemeColor)
                    {
                        Console.WriteLine($"Worksheet: {sheet.Name}, Cell: {cell.Name}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
