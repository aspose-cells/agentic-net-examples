// Title: Load an Excel workbook and preserve original cell formatting when saving with Aspose.Cells for .NET
// Common Searches: how to keep original cell formatting when saving an Excel file with Aspose.Cells in C# | Aspose.Cells .NET iterate through every cell to read its style | save workbook without applying default Excel theme using Aspose.Cells | reapply unchanged cell style after processing with Aspose.Cells C# | check if Excel theme can be disabled programmatically in Aspose.Cells
// Tags: preserve cell formatting Aspose.Cells .NET | cell style iteration Aspose.Cells C# | save workbook without theme Aspose.Cells | read and set cell style Aspose.Cells | disable Excel theme workaround Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook with Aspose.Cells, iterates through each worksheet and every used cell, retrieves each cell's style, reassigns the unchanged style back to the cell, and saves the workbook to a new file, ensuring that the original formatting is retained without applying any theme transformations.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells
                foreach (Cell cell in cells)
                {
                    // Retrieve the cell's style (placeholder for any future processing)
                    Style style = cell.GetStyle();

                    // Example: you can modify style properties here if needed
                    // For now, no theme‑related processing is performed because
                    // the current Aspose.Cells version does not expose Theme APIs.
                    
                    // Apply the (unchanged) style back to the cell
                    cell.SetStyle(style);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
