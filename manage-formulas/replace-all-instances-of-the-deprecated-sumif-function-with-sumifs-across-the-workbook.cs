// Title: C# – Replace all SUMIF formulas with SUMIFS using Aspose.Cells for .NET
// Description: Loads an Excel workbook (creates a sample file if missing), uses Workbook.Replace to change each SUMIF function to SUMIFS, recalculates all formulas, and saves the updated workbook.
// Keywords: Aspose.Cells | C# Excel automation | SUMIF to SUMIFS conversion | Workbook.Replace method | recalculate formulas | .NET Excel library | bulk formula update | Excel 2023 compatibility | US developers | European developers
// Common Searches: replace SUMIF with SUMIFS Aspose.Cells C# | bulk update Excel formulas .NET | convert legacy SUMIF to SUMIFS programmatically | Aspose.Cells replace text in formulas | recalculate workbook after formula replace
// Developer Intent: Swap every SUMIF occurrence for SUMIFS in the loaded workbook.
// Use Cases: Upgrade legacy spreadsheets to the newer SUMIFS syntax before sharing. | Automate mass formula migration across multiple workbooks in a CI/CD pipeline. | Guarantee calculation accuracy after a bulk function replacement. | Prepare Excel files for compatibility with Excel 365 and later versions.
// AI Prompts: Generate C# code using Aspose.Cells that finds and replaces SUMIF with SUMIFS in all worksheets, then recalculates and logs each change. | Show how to replace several deprecated functions (e.g., SUMIF, COUNTIF) in a single pass with Aspose.Cells' Replace method. | Provide robust error‑handling for loading, modifying, and saving an Excel file while performing formula replacements using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook (creates a sample file if missing), uses Workbook.Replace to change each SUMIF function to SUMIFS, recalculates all formulas, and saves the updated workbook.
    public class ReplaceSumIfWithSumIfs
    {
        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists; create a simple workbook if it doesn't
                if (!File.Exists(inputPath))
                {
                    Workbook tempWb = new Workbook();
                    Worksheet ws = tempWb.Worksheets[0];
                    ws.Cells["A1"].Formula = "=SUMIF(B1:B5, \">10\", C1:C5)";
                    tempWb.Save(inputPath);
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Replace deprecated SUMIF with SUMIFS in all formulas
                workbook.Replace("SUMIF", "SUMIFS");

                // Recalculate formulas to reflect changes
                workbook.CalculateFormula();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ReplaceSumIfWithSumIfs.Run();
        }
    }
}
