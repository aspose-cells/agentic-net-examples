// Title: Find and list Excel cells that use volatile functions with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to scan a workbook and return the addresses of all cells whose formulas contain volatile functions such as NOW, TODAY, RAND, OFFSET, etc. | Extend the example to export the identified volatile‑function cell addresses to a CSV file, including worksheet name and cell reference. | Create a reusable C# method that accepts a Worksheet object and a collection of volatile function names, and returns a list of matching cell names using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to detect formulas with volatile functions like NOW or OFFSET | list cell references that contain volatile Excel functions using Aspose.Cells .NET | C# code to enumerate cells with volatile formulas in an Excel workbook with Aspose.Cells | find volatile function usage in Excel file programmatically with Aspose.Cells for .NET | extract addresses of cells using RAND or TODAY in a workbook via Aspose.Cells C#
// Tags: detect volatile functions Aspose.Cells C# | enumerate cells with volatile formulas .NET | locate volatile formula cells in Excel workbook | scan workbook for volatile formulas using Aspose.Cells | identify volatile function usage in Excel file C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsVolatileFinder
{
    // The sample loads an input.xlsx workbook, enables the calculation chain, iterates through each worksheet and its used cells, checks each formula for any of the predefined volatile functions (NOW, TODAY, RAND, OFFSET, etc.), records the sheet name and cell address for matches, prints the list of volatile‑function cells, and saves the workbook as output.xlsx.
    class Program
    {
        // List of common Excel volatile functions
        private static readonly string[] VolatileFunctions = new[]
        {
            "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET()", "INDIRECT()", "INFO()", "CELL()",
            "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL"
        };

        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(inputPath);

                // Enable calculation chain and recalculate formulas
                workbook.Settings.FormulaSettings.EnableCalculationChain = true;
                workbook.CalculateFormula();

                // Collect addresses of cells that contain volatile functions
                List<string> volatileCellAddresses = new List<string>();

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all used cells in the worksheet
                    foreach (Cell cell in cells)
                    {
                        // Check if the cell contains a formula
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula ?? string.Empty;

                            // Simple detection: look for any volatile function name in the formula (case‑insensitive)
                            foreach (string volatileFunc in VolatileFunctions)
                            {
                                if (formula.IndexOf(volatileFunc, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    // Include sheet name for clarity
                                    volatileCellAddresses.Add($"{sheet.Name}!{cell.Name}");
                                    break; // No need to check other functions for this cell
                                }
                            }
                        }
                    }
                }

                // Output the results
                Console.WriteLine("Cells containing volatile functions:");
                foreach (string address in volatileCellAddresses)
                {
                    Console.WriteLine(address);
                }

                // Save the workbook (lifecycle rule: save)
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
