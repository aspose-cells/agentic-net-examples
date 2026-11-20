// Title: How to replace themed cell colors with explicit RGB values in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to strip theme formatting from Excel cells while preserving colors | convert Excel theme-based cell colors to fixed RGB values using Aspose.Cells .NET | programmatically write explicit foreground and background colors to each cell in a workbook with Aspose.Cells | iterate over all worksheets and cells to resolve theme colors in Aspose.Cells | save an XLSX file without any theme references using Aspose.Cells for .NET
// Tags: resolve theme colors to RGB Aspose.Cells | explicit cell style color assignment C# | remove workbook theme references programmatically | iterate worksheets cells set style colors | workaround missing IsThemeEnabled property

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeResolver
{
    // The example loads an XLSX file with Aspose.Cells, walks through every worksheet and cell, reads each cell’s resolved ForegroundColor and BackgroundColor, reassigns those same RGB values back to the style, thereby removing any theme dependencies, and saves the workbook as a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // NOTE: In recent Aspose.Cells versions the IsThemeEnabled property has been removed.
                // The following loop forces the resolved RGB values to be written back to each cell,
                // effectively eliminating any theme references without needing to toggle the theme flag.

                // Iterate through all worksheets and cells to make the resolved colors explicit
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    foreach (Cell cell in cells)
                    {
                        // Get the current style
                        Style style = cell.GetStyle();

                        // Retrieve the resolved foreground and background colors
                        Color foreColor = style.ForegroundColor;
                        Color backColor = style.BackgroundColor;

                        // Re‑assign the same colors to the style.
                        // This writes the explicit RGB values into the cell style,
                        // removing any theme reference.
                        style.ForegroundColor = foreColor;
                        style.BackgroundColor = backColor;

                        // Apply the updated style back to the cell
                        cell.SetStyle(style);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors to prevent the application from crashing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
