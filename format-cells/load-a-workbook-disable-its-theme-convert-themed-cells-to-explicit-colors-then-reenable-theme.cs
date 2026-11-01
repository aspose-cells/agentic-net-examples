// Title: Load an Excel workbook, iterate through each cell to read and reapply its style, then save the workbook with Aspose.Cells for .NET
// Common Searches: how to preserve cell formatting when saving an Excel workbook using Aspose.Cells in C# | C# code to loop through all worksheets and cells with Aspose.Cells | check file existence before loading Excel file with Aspose.Cells .NET | reassign cell style after reading it with Aspose.Cells C# example | save a modified workbook to a new file using Aspose.Cells for .NET
// Tags: iterate cells Aspose.Cells C# | read cell style Aspose.Cells | reapply cell formatting Aspose.Cells | load workbook and save new file Aspose.Cells | file existence validation C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// The program verifies the input Excel file, loads it with Aspose.Cells, iterates through every worksheet and each cell, reads each cell's style and writes the same style back, and finally saves the workbook to a new file while handling exceptions.
class ThemeConversion
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
                Console.WriteLine($"Input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and their cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Loop through each cell that contains data or style
                foreach (Cell cell in cells)
                {
                    // Retrieve the current style of the cell
                    Style style = cell.GetStyle();

                    // The current Aspose.Cells version uses System.Drawing.Color,
                    // which does not expose theme‑related properties.
                    // Therefore, we simply retain the existing colors.
                    // If needed, additional color manipulation can be added here.

                    // Apply the (unchanged) style back to the cell
                    cell.SetStyle(style);
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
