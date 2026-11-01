// Title: Read each cell’s style and prepare for theme‑color detection with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells read cell style and theme color information from an Excel workbook | how to iterate all used cells and get their Style object using Aspose.Cells .NET | detect whether an Excel workbook uses theme colors with Aspose.Cells in C# | fallback to direct RGB color when theme color is unavailable using Aspose.Cells | access cell formatting details programmatically with Aspose.Cells for .NET
// Tags: read cell style Aspose.Cells | iterate workbook cells Aspose.Cells .NET | detect theme color usage Aspose.Cells | fallback to RGB formatting Aspose.Cells | access Excel cell formatting programmatically | style object handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// The example loads an Excel workbook with Aspose.Cells, walks through every worksheet and each used cell, retrieves the cell's Style object, re‑applies the unchanged style, and saves the workbook. This pattern shows how to access cell formatting and can be extended to detect theme‑color usage and replace it with direct RGB values when needed.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Loop through used cells
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        Style style = cell.GetStyle();

                        // If needed, additional style processing can be added here.
                        // Current implementation retains existing formatting.

                        // Apply the (unchanged) style back to the cell
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
