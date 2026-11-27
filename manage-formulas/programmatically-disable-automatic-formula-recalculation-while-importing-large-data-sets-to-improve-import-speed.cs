// Title: How to disable automatic formula recalculation in Aspose.Cells for .NET to speed up bulk data imports
// Common Searches: Aspose.Cells disable auto calculation before loading 100k rows C# | set workbook calculation mode to manual for large data import Aspose.Cells .NET | improve Excel export performance by turning off formula recalculation Aspose.Cells | C# bulk insert data into Aspose.Cells workbook without triggering recalculation | manual calculation mode Aspose.Cells example for high‑volume data
// Tags: manual calculation mode Aspose.Cells | bulk data import performance .NET | disable auto formula recalculation Aspose.Cells | large dataset workbook creation C# | optimize Excel export speed Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, sets its calculation mode to Manual to prevent automatic formula evaluation, fills the first worksheet with 100,000 rows × 10 columns of sample values, optionally calls CalculateFormula after the bulk load, ensures the output directory exists, and saves the workbook as LargeData.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: import a large data set (e.g., 100,000 rows × 10 columns)
            int totalRows = 100000;
            int totalCols = 10;

            // Populate cells with sample data
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    sheet.Cells[row, col].PutValue(row * col);
                }
            }

            // Optionally recalculate formulas manually (if any)
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "LargeData.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
