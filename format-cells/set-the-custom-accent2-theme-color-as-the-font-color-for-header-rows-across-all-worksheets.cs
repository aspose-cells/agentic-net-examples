// Title: Set the Accent2 theme font color for header rows in every worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change header row font color to Accent2 across all sheets | How to apply a custom theme color to the first row of each worksheet in a .NET Excel file | Programmatically set Excel header font to blue Accent2 using Aspose.Cells | Loop through worksheets and set header font color with Aspose.Cells .NET | Set theme-based font color for Excel table headers in C#
// Tags: accent2 font color for Excel headers Aspose.Cells | format header rows across all worksheets C# | apply custom theme color to cell font Aspose.Cells | bulk header row styling .NET Excel | set theme-based font color programmatically C#

using Aspose.Cells;
using System;
using System.IO;
using System.Drawing;

// The program loads an existing workbook, iterates through each worksheet, and changes the font color of every cell in the first row to the Accent2 theme blue (RGB 0,112,192) before saving the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Assume the header row is the first row (index 0)
                Row headerRow = sheet.Cells.Rows[0];

                // Apply a theme-like Accent2 color to each cell in the header row
                foreach (Cell cell in headerRow)
                {
                    // Get the current style of the cell
                    Style style = cell.GetStyle();

                    // Set the font color to an approximate Accent2 theme color
                    style.Font.Color = Color.FromArgb(0, 112, 192); // typical Accent2 blue

                    // Apply the modified style back to the cell
                    cell.SetStyle(style);
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
