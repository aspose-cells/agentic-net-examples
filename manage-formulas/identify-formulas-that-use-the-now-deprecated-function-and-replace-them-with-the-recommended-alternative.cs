// Title: Replace deprecated Excel functions with Aspose.Cells for .NET across all worksheets
// Description: Creates a sample workbook if needed, loads an existing Excel file, scans every cell for a formula that uses a deprecated function (e.g., OLDFUNC), replaces it with the recommended alternative (e.g., NEWFUNC) using a case‑insensitive replace, recalculates all formulas, and saves the updated file.
// Keywords: Aspose.Cells replace deprecated function | update Excel formulas .NET | find and replace Excel function | recalculate workbook after formula change | OLDFUNC NEWFUNC Aspose.Cells | bulk formula update C# | Excel migration legacy functions
// Common Searches: how to replace a deprecated Excel function with Aspose.Cells | Aspose.Cells iterate cells to modify formulas | recalculate workbook after formula replacement .NET | bulk replace Excel functions using C# | case insensitive formula replace Aspose.Cells
// Developer Intent: Automatically locate every occurrence of a deprecated Excel function in a workbook, substitute it with the supported alternative, recalculate the sheet, and save the result.
// Use Cases: Modernize legacy spreadsheets when upgrading to newer Excel versions. | Sanitize user‑uploaded workbooks by swapping unsupported functions before processing. | Generate a placeholder workbook when the source file is missing, then apply function replacement.
// AI Prompts: Generate C# code with Aspose.Cells that searches all formulas for 'OLD_FUNC', replaces them with 'NEW_FUNC' case‑insensitively, and recalculates the workbook. | Provide an Aspose.Cells snippet that logs each cell whose formula was changed during the replacement process. | Explain how to perform a bulk, case‑insensitive replacement of deprecated functions in Excel formulas using Aspose.Cells and ensure formulas are re‑evaluated.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a sample workbook if needed, loads an existing Excel file, scans every cell for a formula that uses a deprecated function (e.g., OLDFUNC), replaces it with the recommended alternative (e.g., NEWFUNC) using a case‑insensitive replace, recalculates all formulas, and saves the updated file.
    public class ReplaceDeprecatedFunctionDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a sample if it does not.
            if (!File.Exists(inputPath))
            {
                var sampleWorkbook = new Workbook();
                sampleWorkbook.Worksheets[0].Cells["A1"].Formula = "OLDFUNC(1,2)";
                sampleWorkbook.Save(inputPath);
                Console.WriteLine($"Sample input file created at '{inputPath}'.");
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            const string deprecatedFunc = "OLDFUNC";
            const string replacementFunc = "NEWFUNC";

            // Replace deprecated functions in all formulas.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula) &&
                        cell.Formula.IndexOf(deprecatedFunc, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string updatedFormula = cell.Formula.Replace(deprecatedFunc, replacementFunc, StringComparison.OrdinalIgnoreCase);
                        cell.Formula = updatedFormula;
                    }
                }
            }

            // Recalculate formulas after replacement.
            workbook.CalculateFormula();

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
