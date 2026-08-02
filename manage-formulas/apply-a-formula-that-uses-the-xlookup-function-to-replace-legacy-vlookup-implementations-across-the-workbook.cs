// Title: Convert VLOOKUP to XLOOKUP across all worksheets using Aspose.Cells for .NET
// Description: Loads an Excel file, scans every worksheet for VLOOKUP formulas, swaps the function name to XLOOKUP, recalculates the workbook, and saves the updated file with Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | .NET | VLOOKUP to XLOOKUP conversion | replace Excel formulas programmatically | batch formula update | search and replace formulas | recalculate workbook | Excel automation
// Common Searches: Aspose.Cells replace VLOOKUP with XLOOKUP C# example | find and update Excel formulas across all sheets .NET | convert legacy VLOOKUP formulas to XLOOKUP programmatically | batch replace VLOOKUP in workbook using Aspose.Cells | recalculate formulas after XLOOKUP conversion
// Developer Intent: Programmatically replace every VLOOKUP formula in a workbook with an XLOOKUP equivalent using Aspose.Cells for .NET.
// Use Cases: Modernize legacy spreadsheets before distribution. | Ensure consistent formula syntax across multiple workbooks. | Automate bulk migration of VLOOKUP to XLOOKUP in enterprise reporting. | Validate results by recalculating after conversion.
// AI Prompts: Generate C# code that parses VLOOKUP arguments and builds accurate XLOOKUP formulas with Aspose.Cells. | Show how to log each formula change, including original VLOOKUP text and the new XLOOKUP expression. | Provide robust error handling for cases where VLOOKUP cannot be directly mapped to XLOOKUP.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel file, scans every worksheet for VLOOKUP formulas, swaps the function name to XLOOKUP, recalculates the workbook, and saves the updated file with Aspose.Cells in C#.
    public class ReplaceVlookupWithXlookup
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
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

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Configure find options to search only in formulas and allow partial matches
                    FindOptions findOptions = new FindOptions
                    {
                        LookInType = LookInType.OnlyFormulas,
                        LookAtType = LookAtType.Contains
                    };

                    // Find the first cell that contains the legacy VLOOKUP function
                    Cell cell = sheet.Cells.Find("VLOOKUP", null, findOptions);

                    // Continue processing while such cells are found
                    while (cell != null)
                    {
                        // Retrieve the original VLOOKUP formula
                        string oldFormula = cell.Formula;

                        // Simple textual conversion: replace the function name.
                        // For a production scenario you would need a more robust parser to
                        // rearrange arguments according to XLOOKUP's signature.
                        string newFormula = oldFormula.Replace("VLOOKUP", "XLOOKUP");

                        // Apply the new formula to the same cell
                        cell.SetFormula(newFormula, new FormulaParseOptions());

                        // Search for the next occurrence in the same worksheet
                        cell = sheet.Cells.Find("VLOOKUP", cell, findOptions);
                    }
                }

                // Recalculate all formulas so that the newly set XLOOKUP formulas are evaluated
                workbook.CalculateFormula();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
