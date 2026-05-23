using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data and a formula
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].Formula = "=SUM(A1:A2)";

            // Set calculation options (multi‑threaded calculation can be enabled if needed)
            CalculationOptions options = new CalculationOptions
            {
                // Enable multi‑threaded calculation if the API supports it
                // (ThreadCount property is not available in the current version)
                // EnableThreadedCalculation = true
            };

            // Calculate all formulas using the specified options
            workbook.CalculateFormula(options);

            // Display the calculated result
            Console.WriteLine("A3 = " + cells["A3"].Value);

            // Save the workbook (lifecycle save rule)
            string outputPath = "LimitedThreadsResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}