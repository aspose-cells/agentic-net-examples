// Title: Load an Excel workbook, iterate every cell to read and reapply its style, then save the file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a .xlsx file with Aspose.Cells, loops through all worksheets and each cell in the used range, retrieves the cell's Style object, assigns the same style back to the cell, and writes the workbook to a new file while handling missing input files and exceptions. | Create a .NET example that demonstrates loading a workbook, enumerating every cell to read its formatting, reapplying the retrieved Style, and saving the modified workbook using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# iterate over all cells in a workbook and reapply style | how to read and set cell style for each cell using Aspose.Cells .NET | load Excel file, loop through used range, and save with Aspose.Cells C# | Aspose.Cells example for reapplying cell formatting after loading workbook | C# Aspose.Cells handling missing input.xlsx file before processing
// Tags: Aspose.Cells iterate cells C# | Aspose.Cells reapply cell style .NET | Aspose.Cells load and save workbook example | Aspose.Cells used range cell enumeration | Aspose.Cells exception handling missing file

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads 'input.xlsx' with Aspose.Cells, iterates through every worksheet and each cell in the used range, retrieves each cell's Style, immediately reassigns the same style back to the cell, and saves the result as 'output.xlsx', including checks for the input file's existence and basic exception handling.
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

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Loop through each cell in the used range
                foreach (Cell cell in cells)
                {
                    // Retrieve the current style
                    Style style = cell.GetStyle();

                    // If needed, modify style properties here.
                    // (Theme‑related properties are not available in the current API version.)

                    // Apply the (potentially modified) style back to the cell
                    cell.SetStyle(style);
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
