// Title: How to evaluate MAXIFS formula accuracy with Aspose.Cells .NET default calculation engine as a baseline for custom cached engine comparison
// Common Searches: baseline MAXIFS result using Aspose.Cells default calculation engine for .NET | compare Aspose.Cells default and custom calculation engine performance for MAXIFS | benchmark MAXIFS formula evaluation with caching in Aspose.Cells .NET | run MAXIFS with Aspose.Cells and capture result for custom engine testing | Aspose.Cells .NET formula calculation accuracy example for MAXIFS
// Tags: MAXIFS formula evaluation Aspose.Cells | default calculation engine benchmarking Aspose.Cells | custom calculation engine caching .NET | Excel formula performance testing Aspose.Cells | Aspose.Cells workbook calculation accuracy

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example creates a workbook, populates columns A and B with numeric values and criteria, inserts a MAXIFS formula in cell C2 to find the maximum value where the criteria are >= 3, calculates the formula using Aspose.Cells' default calculation engine, outputs the result, and saves the workbook as MaxIfsComparison.xlsx, providing a baseline for later comparison with a custom cached calculation engine.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for MAXIFS
                sheet.Cells["A1"].PutValue("Values");
                sheet.Cells["B1"].PutValue("Criteria");

                double[] values = { 10, 20, 30, 40, 50 };
                int[] criteria = { 1, 2, 3, 4, 5 };

                for (int i = 0; i < values.Length; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(values[i]);   // A2:A6
                    sheet.Cells[i + 1, 1].PutValue(criteria[i]); // B2:B6
                }

                // Insert MAXIFS formula that selects the maximum value where criteria >= 3
                sheet.Cells["C2"].Formula = "MAXIFS(A2:A6,B2:B6,\">=3\")";

                // Calculate using the default calculation engine
                workbook.CalculateFormula();
                double defaultResult = sheet.Cells["C2"].DoubleValue;

                // Output the result
                Console.WriteLine($"Default engine result: {defaultResult}");

                // Save the workbook for verification
                string outputPath = "MaxIfsComparison.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
