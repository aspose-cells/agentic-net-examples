// Title: Set the Accent2 theme color as the font color for header rows in every worksheet with Aspose.Cells for .NET
// Common Searches: how to change header row font color to Accent2 using Aspose.Cells C# | Aspose.Cells set theme accent color for first row of each sheet | apply bold blue Accent2 font to Excel headers in .NET | iterate all worksheets and style header cells with Aspose.Cells | C# code to apply custom theme color to Excel header rows using Aspose.Cells
// Tags: header row font styling Aspose.Cells | Accent2 theme color Excel .NET | apply bold font to first row Aspose.Cells | format all worksheets header cells C# | custom theme color for Excel headers Aspose

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, loops through each worksheet, and for the first (header) row of every non‑empty sheet sets the font color to a blue shade representing the Accent2 theme and makes the text bold, then saves the updated workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                try
                {
                    // Skip empty worksheets
                    if (worksheet.Cells.MaxRow < 0)
                        continue;

                    // Determine the last used column in the first row (header)
                    int lastColumn = worksheet.Cells.MaxColumn;

                    // Apply a theme-like accent color and bold font to each header cell
                    for (int col = 0; col <= lastColumn; col++)
                    {
                        Cell cell = worksheet.Cells[0, col];
                        Style style = cell.GetStyle();

                        // Use a typical Accent2 color (e.g., a shade of blue)
                        style.Font.Color = Color.FromArgb(0, 112, 192);
                        style.Font.IsBold = true;

                        cell.SetStyle(style);
                    }
                }
                catch (Exception exWorksheet)
                {
                    Console.WriteLine($"Error processing worksheet '{worksheet.Name}': {exWorksheet.Message}");
                }
            }

            // Save the modified workbook to the output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
