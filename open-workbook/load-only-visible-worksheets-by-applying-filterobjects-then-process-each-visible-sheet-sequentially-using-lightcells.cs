// Title: Load an Excel workbook and iterate only visible worksheets to process each cell with Aspose.Cells in C#
// Common Searches: how to skip hidden worksheets when reading cells with Aspose.Cells C# | process cells only on visible sheets using Aspose.Cells .NET example | iterate through visible worksheets in an Excel file with Aspose.Cells C# code | Aspose.Cells C# read data from visible worksheets only | filter out hidden sheets before cell iteration in Aspose.Cells
// Tags: iterate visible worksheets Aspose.Cells | cell processing on visible sheets .NET | skip hidden worksheets Aspose.Cells | read cells from visible worksheets C# | Aspose.Cells workbook iteration visible sheets

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input file, loads it into an Aspose.Cells Workbook, loops through all worksheets, skips any that are not visible, and then iterates over each cell up to the maximum data row and column, calling a custom ProcessCell method before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Process only worksheets that are visible
                if (!sheet.IsVisible)
                    continue;

                try
                {
                    // Process each cell in the visible worksheet
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            ProcessCell(cell, row, col);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing sheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            // Save the workbook after processing (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Example cell processing method (customize as needed)
    private static void ProcessCell(Cell cell, int row, int column)
    {
        // Read the cell value (add your own logic here)
        object value = cell.Value;
        // TODO: implement required processing for each cell
    }
}
