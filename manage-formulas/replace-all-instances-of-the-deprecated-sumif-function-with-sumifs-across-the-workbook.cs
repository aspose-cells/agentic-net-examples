// Title: C# – Replace SUMIF with SUMIFS in an Excel workbook using Aspose.Cells
// Description: Loads an Excel file, uses Workbook.Replace to change every "SUMIF" occurrence to "SUMIFS" in formulas and cell text, recalculates the workbook, and saves the updated file.
// Keywords: Aspose.Cells | C# | Workbook.Replace | SUMIF | SUMIFS | replace Excel function | bulk formula update | recalculate formulas | Excel automation | legacy workbook conversion
// Common Searches: Aspose.Cells replace SUMIF with SUMIFS | C# replace Excel function across workbook | bulk update Excel formulas programmatically | recalculate formulas after text replace Aspose.Cells | convert legacy SUMIF to SUMIFS using .NET
// Developer Intent: Programmatically replace all instances of the SUMIF function with SUMIFS in every worksheet, recalculate formulas, and save the updated workbook.
// Use Cases: Upgrade legacy spreadsheets that contain deprecated SUMIF before opening them in newer Excel versions. | Automate mass conversion of Excel files in a migration or CI/CD pipeline. | Maintain data integrity by forcing formula recalculation after function replacement.
// AI Prompts: Write C# code using Aspose.Cells to replace SUMIF with SUMIFS in all cells and recalculate the workbook. | Show how to add robust error handling for missing input files and log the number of replacements performed. | Explain the difference between replacing text in formulas versus plain cell values and why CalculateFormula is required afterward. | Generate a unit test that verifies SUMIF was replaced by SUMIFS in a sample workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel file, uses Workbook.Replace to change every "SUMIF" occurrence to "SUMIFS" in formulas and cell text, recalculates the workbook, and saves the updated file.
    public class ReplaceSumIfWithSumIfs
    {
        public static void Run()
        {
            try
            {
                // Path to the input workbook
                string inputPath = "InputWorkbook.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Replace deprecated SUMIF with SUMIFS in all formulas and text
                int replacedCount = workbook.Replace("SUMIF", "SUMIFS");
                Console.WriteLine($"Replaced {replacedCount} occurrences of SUMIF with SUMIFS.");

                // Recalculate formulas to evaluate the new SUMIFS functions
                workbook.CalculateFormula();

                // Save the updated workbook
                string outputPath = "OutputWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ReplaceSumIfWithSumIfs.Run();
        }
    }
}
