// Title: C# – Scan Excel workbook for volatile formulas with Aspose.Cells
// Description: Loads an .xlsx file using Aspose.Cells, iterates through every worksheet and cell, detects formulas that contain volatile Excel functions (NOW, TODAY, RAND, RANDBETWEEN, OFFSET, INDIRECT, INFO, CELL), logs the sheet name, cell address and formula, and optionally saves the workbook. Perfect for performance‑impact analysis.
// Keywords: Aspose.Cells | C# | .NET | volatile Excel functions | formula scanner | performance optimization | NOW function | OFFSET function | Excel recalculation | scan workbook | detect volatile formulas
// Common Searches: Aspose.Cells find volatile formulas | C# list cells with NOW or OFFSET | detect Excel volatile functions programmatically | performance audit Excel formulas .NET | scan workbook for volatile functions using Aspose
// Developer Intent: Identify every formula that uses a volatile Excel function and capture its location for optimization.
// Use Cases: Generate a performance‑audit report of all volatile formulas in a workbook. | Flag or replace volatile functions during automated workbook cleanup. | Integrate the scan into a CI pipeline to enforce formula‑performance standards. | Export the list of volatile formulas to CSV/JSON for further analysis.
// AI Prompts: Create a method that returns a List of objects containing worksheet name, cell address, and formula for all volatile functions found. | Modify the sample to write the volatile‑formula report to a CSV file instead of the console. | Add a parameter that accepts a custom volatile‑function list and logs matches with line numbers. | Implement parallel processing to accelerate scanning of very large workbooks.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an .xlsx file using Aspose.Cells, iterates through every worksheet and cell, detects formulas that contain volatile Excel functions (NOW, TODAY, RAND, RANDBETWEEN, OFFSET, INDIRECT, INFO, CELL), logs the sheet name, cell address and formula, and optionally saves the workbook. Perfect for performance‑impact analysis.
    public class VolatileFormulaScanner
    {
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
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the existing workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // List of known volatile Excel functions
            List<string> volatileFunctions = new List<string>
            {
                "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET()", "INDIRECT()", "INFO()", "CELL()",
                "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL"
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain data or formulas
                foreach (Cell cell in cells)
                {
                    // Check if the cell has a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        string formula = cell.Formula;

                        // Detect presence of any volatile function (case‑insensitive)
                        foreach (string volatileFunc in volatileFunctions)
                        {
                            if (formula.IndexOf(volatileFunc, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                Console.WriteLine($"Worksheet: {sheet.Name}, Cell: {cell.Name}, Formula: {formula}");
                                break; // No need to check other volatile functions for this cell
                            }
                        }
                    }
                }
            }

            try
            {
                // Save the workbook (optional, demonstrates lifecycle usage)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
