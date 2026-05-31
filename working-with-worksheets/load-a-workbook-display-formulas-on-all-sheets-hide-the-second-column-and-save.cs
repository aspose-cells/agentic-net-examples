using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ShowFormulasAndHideColumn
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Process each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Show formulas instead of calculated values
                    sheet.ShowFormulas = true;

                    // Hide the second column (B)
                    sheet.Cells.HideColumn(1);
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