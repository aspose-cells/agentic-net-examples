using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a large range with sample numeric data
            for (int i = 0; i < 10000; i++)
            {
                cells[i, 0].PutValue(i + 1);
            }

            // Add a formula that sums the entire column
            cells[0, 1].Formula = $"=SUM(A1:A{cells.MaxDataRow + 1})";

            // Calculate all formulas in the workbook (default options include multi‑threaded calculation when supported)
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "MultiThreadedCalculation.xlsx";

            // Save the resulting workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}