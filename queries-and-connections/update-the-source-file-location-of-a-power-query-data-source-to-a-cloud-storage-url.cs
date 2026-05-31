using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsExamples
{
    public class UpdatePowerQuerySourceFileLocation
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook containing Power Query connections
                Workbook workbook = new Workbook(inputPath);

                // Define old path segment and new cloud storage URL
                string oldPathSegment = @"C:\Data\";
                string newCloudUrl = "https://cloudstorage.example.com/data/";

                // Iterate through all Power Query formulas
                foreach (PowerQueryFormula formula in workbook.DataMashup.PowerQueryFormulas)
                {
                    foreach (PowerQueryFormulaItem item in formula.PowerQueryFormulaItems)
                    {
                        // Replace old path with new URL if present
                        if (!string.IsNullOrEmpty(item.Value) && item.Value.Contains(oldPathSegment))
                        {
                            string updatedValue = item.Value.Replace(oldPathSegment, newCloudUrl);
                            item.Value = updatedValue;
                            Console.WriteLine($"Updated Power Query item: {updatedValue}");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}