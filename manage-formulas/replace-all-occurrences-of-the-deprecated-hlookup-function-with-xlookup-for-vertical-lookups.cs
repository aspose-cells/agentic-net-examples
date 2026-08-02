// Title: Replace HLOOKUP with XLOOKUP in Excel using Aspose.Cells for .NET
// Description: Loads a workbook, scans every worksheet and cell, finds formulas that contain HLOOKUP, swaps the function name to XLOOKUP (case‑insensitive), recalculates all formulas, and saves the updated file. Perfect for bulk migration of legacy lookups to the modern XLOOKUP function.
// Keywords: Aspose.Cells | C# | XLOOKUP migration | replace HLOOKUP | Excel formula update | bulk formula replacement | programmatic Excel | Excel 365 compatibility | workbook recalculation | Excel automation .NET
// Common Searches: Aspose.Cells replace HLOOKUP with XLOOKUP | C# code to convert HLOOKUP to XLOOKUP | bulk update Excel formulas using Aspose.Cells | programmatically change Excel lookup functions .NET | migrate legacy HLOOKUP spreadsheets to XLOOKUP
// Developer Intent: Programmatically replace every HLOOKUP formula with an XLOOKUP equivalent in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Modernize legacy reporting workbooks before distribution. | Ensure compatibility with Excel 365 by upgrading lookup functions. | Automate large‑scale formula updates across multiple worksheets. | Recalculate the workbook after changes to verify results.
// AI Prompts: Generate C# code with Aspose.Cells that scans all cells in a workbook and replaces HLOOKUP with XLOOKUP, preserving arguments and recalculating the workbook. | Provide a robust method to convert HLOOKUP formulas to XLOOKUP, handling case‑insensitive matches and saving the updated file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, scans every worksheet and cell, finds formulas that contain HLOOKUP, swaps the function name to XLOOKUP (case‑insensitive), recalculates all formulas, and saves the updated file. Perfect for bulk migration of legacy lookups to the modern XLOOKUP function.
    public class ReplaceHlookupWithXlookup
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit the iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only cells that contain a formula with HLOOKUP
                        if (!string.IsNullOrEmpty(cell.Formula) &&
                            cell.Formula.IndexOf("HLOOKUP", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Convert the HLOOKUP formula to an XLOOKUP formula
                            string updatedFormula = ConvertHlookupToXlookup(cell.Formula);
                            cell.Formula = updatedFormula;
                        }
                    }
                }
            }

            // Recalculate all formulas after the replacement
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }

        private static string ConvertHlookupToXlookup(string hlookupFormula)
        {
            // Replace the function name while preserving the rest of the formula
            return hlookupFormula.Replace("HLOOKUP", "XLOOKUP", StringComparison.OrdinalIgnoreCase);
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                ReplaceHlookupWithXlookup.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
