using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ReplaceDeprecatedFormulasDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Loop through all used cells
                    foreach (Cell cell in cells)
                    {
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            // Detect deprecated function "INDIRECT"
                            if (formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Replace with "INDEX"
                                string updatedFormula = formula.Replace("INDIRECT", "INDEX", StringComparison.OrdinalIgnoreCase);
                                cell.Formula = updatedFormula;

                                // Clear cached value to force recalculation
                                cell.Value = null;
                            }
                        }
                    }
                }

                // Recalculate all formulas after modifications
                workbook.CalculateFormula();

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File error: {fnfEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}