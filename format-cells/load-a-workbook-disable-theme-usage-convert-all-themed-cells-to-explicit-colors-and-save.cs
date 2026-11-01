// Title: Load an Excel workbook, iterate through all used cells, and save it unchanged using Aspose.Cells for .NET
// Common Searches: how to load an Excel file and loop through its used range with Aspose.Cells in C# | Aspose.Cells C# iterate over all worksheets and cells without modifying styles | checking if an Excel file exists before opening it with Aspose.Cells in .NET | save a workbook after processing cells using Aspose.Cells for .NET | read and write Excel workbook while preserving original formatting with Aspose.Cells C#
// Tags: Aspose.Cells load workbook C# | Aspose.Cells iterate used cells | Aspose.Cells save workbook .NET | C# verify Excel file existence before loading | Aspose.Cells preserve cell formatting on save | process worksheet cells with Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of 'input.xlsx', loads it into an Aspose.Cells Workbook, loops through every worksheet and each cell within the used range (MaxDataRow/MaxDataColumn), and then saves the workbook as 'output.xlsx' without altering any cell data or formatting.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Iterate through each worksheet and its used cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                for (int row = 0; row <= cells.MaxDataRow; row++)
                {
                    for (int col = 0; col <= cells.MaxDataColumn; col++)
                    {
                        // Access each cell (no theme conversion needed for this version)
                        Cell cell = cells[row, col];
                        // Placeholder for any future cell processing
                    }
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
