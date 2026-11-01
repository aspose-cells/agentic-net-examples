// Title: How to set every cell’s font size to 12 points in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change font size for all cells in a workbook | set uniform 12‑point font across entire Excel file with Aspose.Cells | iterate worksheets and apply font size to each cell using Aspose.Cells .NET | bulk update cell style font size Aspose.Cells C# example
// Tags: bulk cell font size update Aspose.Cells | uniform font size Excel workbook C# | cell style modification Aspose.Cells | apply font size to all cells programmatically | Excel workbook cell formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, loops through every worksheet and each cell, changes each cell's font size to 12 points by updating its Style, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    // Example modification: set font size to 12 points
                    Style style = cell.GetStyle();
                    style.Font.Size = 12;
                    cell.SetStyle(style);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
