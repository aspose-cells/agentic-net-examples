// Title: Verify English function names and comma separators in cell formulas with Aspose.Cells for .NET
// Description: This .NET example shows how to set a workbook's region (e.g., Germany), assign formulas using both the standard English syntax (Formula) and a localized syntax (FormulaLocal), and then validate each formula. The verification extracts the function name, maps it to the standard English name via GetStandardFunctionName, checks that arguments are separated by commas, and outputs the results before saving the workbook.
// Keywords: Aspose.Cells | .NET | formula verification | English function name | comma separator | FormulaLocal | GetStandardFunctionName | globalization | workbook region | German locale | Excel formula validation
// Common Searches: Aspose.Cells check if formula uses English function name | Validate comma separators in Aspose.Cells formulas | How to use GetStandardFunctionName in Aspose.Cells | Formula vs FormulaLocal Aspose.Cells example | Set workbook region for localized formulas Aspose.Cells | Detect non‑English Excel functions with Aspose.Cells
// Developer Intent: Ensure that a cell's formula is expressed with an English function name and commas as argument delimiters.
// Use Cases: Audit workbooks created in non‑English locales to confirm formulas conform to the standard English syntax required by downstream processing. | Automatically convert or flag formulas that use localized function names or semicolon delimiters before exporting to CSV or other systems. | Integrate formula language validation into CI pipelines to enforce compliance with corporate Excel standards.
// AI Prompts: Generate a C# method that receives a Cell and its Workbook and returns true only when the formula uses an English function name and commas as separators, using Aspose.Cells APIs. | Write code that scans all worksheets in a workbook, verifies each formula's language and delimiter, and returns a list of cells that fail the check. | Provide a logging routine that records the cell address, original formula, detected language status, and delimiter type for every formula in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaVerification
{
    // This .NET example shows how to set a workbook's region (e.g., Germany), assign formulas using both the standard English syntax (Formula) and a localized syntax (FormulaLocal), and then validate each formula. The verification extracts the function name, maps it to the standard English name via GetStandardFunctionName, checks that arguments are separated by commas, and outputs the results before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set the workbook region to Germany so that German localized formulas are accepted
                workbook.Settings.Region = CountryCode.Germany;

                // Example 1: Set a formula using standard English function name (Formula property always uses English)
                Cell englishCell = sheet.Cells["A1"];
                englishCell.Formula = "=SUM(1,2,3)"; // English function with commas

                // Example 2: Set a formula using a localized (German) function name
                // This uses the localized property which respects the workbook's region settings
                Cell localizedCell = sheet.Cells["B1"];
                localizedCell.FormulaLocal = "=SUMME(1;2;3)"; // German function with semicolon

                // Verify each cell
                VerifyFormula(englishCell, workbook);
                VerifyFormula(localizedCell, workbook);

                // Save the workbook (lifecycle rule)
                string outputPath = "FormulaVerificationResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <param name="cell">The cell to verify.</param>
        /// <param name="workbook">The workbook containing the cell (needed for globalization settings).</param>
        static void VerifyFormula(Cell cell, Workbook workbook)
        {
            try
            {
                // Retrieve the formula in standard (English) format
                string standardFormula = cell.Formula;          // English function name, commas
                // Retrieve the formula in localized format (if any)
                string localFormula = cell.FormulaLocal;        // May be empty if not set

                // Determine which representation to validate
                string formulaToCheck = string.IsNullOrEmpty(standardFormula) ? localFormula : standardFormula;

                if (string.IsNullOrEmpty(formulaToCheck))
                {
                    Console.WriteLine($"Cell {cell.Name} does not contain a formula.");
                    return;
                }

                // Extract the function name (text before the first '(')
                int openParenIndex = formulaToCheck.IndexOf('(');
                if (openParenIndex <= 0)
                {
                    Console.WriteLine($"Cell {cell.Name}: Unable to parse function name.");
                    return;
                }

                string functionName = formulaToCheck.Substring(0, openParenIndex)
                                                   .TrimStart('=')
                                                   .Trim();

                // Get the standard (English) name for the extracted function name
                string standardName = workbook.Settings.GlobalizationSettings.GetStandardFunctionName(functionName);
                bool isEnglishFunction = string.Equals(functionName, standardName, StringComparison.OrdinalIgnoreCase);

                // Extract the argument list (between '(' and ')')
                int closeParenIndex = formulaToCheck.LastIndexOf(')');
                string arguments = formulaToCheck.Substring(openParenIndex + 1,
                                                            closeParenIndex - openParenIndex - 1);

                bool usesCommas = arguments.Contains(",");

                // Output verification result
                Console.WriteLine($"Cell {cell.Name}:");
                Console.WriteLine($"  Formula: {formulaToCheck}");
                Console.WriteLine($"  Function name '{functionName}' is English: {isEnglishFunction}");
                Console.WriteLine($"  Arguments use commas: {usesCommas}");
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error verifying cell {cell?.Name}: {ex.Message}");
            }
        }
    }
}
