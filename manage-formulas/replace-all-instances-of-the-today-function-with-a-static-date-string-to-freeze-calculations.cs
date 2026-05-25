using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class FreezeTodayFunction
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the static date string to replace TODAY()
            string staticDateString = DateTime.Today.ToString("yyyy-MM-dd");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Check if the cell contains a formula that uses TODAY()
                    if (cell.IsFormula && cell.Formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace the formula with the static date string value
                        cell.PutValue(staticDateString);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}