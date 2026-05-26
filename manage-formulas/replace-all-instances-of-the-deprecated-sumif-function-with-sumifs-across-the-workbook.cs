using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ReplaceSumIfWithSumIfs
    {
        // Entry point required by the runtime
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Replace all occurrences of the deprecated SUMIF function with SUMIFS
                int replacedCount = workbook.Replace("SUMIF(", "SUMIFS(");
                Console.WriteLine($"Replaced {replacedCount} occurrences of SUMIF with SUMIFS.");

                // Recalculate all formulas to reflect the changes
                workbook.CalculateFormula();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}