// Title: C# code to scan an Excel workbook for volatile formulas and write their sheet‑cell locations to a log using Aspose.Cells
// Common Searches: how can I find volatile functions like NOW or OFFSET in an Excel file with Aspose.Cells C# | C# example to list cells containing volatile formulas in a workbook | Aspose.Cells scan workbook for performance‑impacting formulas and export results | write a log of volatile Excel formulas with sheet and cell reference using .NET | detect and record volatile Excel functions programmatically in C#
// Tags: scan workbook for volatile formulas Aspose.Cells | log sheet!cell locations of volatile Excel functions C# | identify performance impacting formulas using Aspose.Cells | extract volatile function references from Excel with .NET | write volatile formula report to .log file Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace VolatileFormulaScanner
{
    // The sample loads an Excel workbook with Aspose.Cells, iterates through every worksheet and cell, checks each formula against a predefined list of volatile functions, records the sheet name and cell address of any matches, prints the findings to the console, and writes them to a .volatile.log file placed beside the original workbook.
    class Program
    {
        // List of known volatile functions in Excel (case‑insensitive)
        private static readonly string[] VolatileFunctions = new[]
        {
            "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET()", "INDIRECT()", "INFO()", "CELL()", "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL"
        };

        static void Main(string[] args)
        {
            // Expect the first argument to be the path of the workbook to analyze.
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the Excel file as an argument.");
                return;
            }

            string workbookPath = args[0];

            // Load the workbook.
            Workbook workbook = new Workbook(workbookPath);

            // Store locations of volatile formulas.
            List<string> volatileFormulaLocations = new List<string>();

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the cells collection.
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain formulas.
                foreach (Cell cell in cells)
                {
                    if (!cell.IsFormula) continue; // Skip non‑formula cells.

                    string formula = cell.Formula; // Formula string as entered (e.g., "=NOW()").

                    // Check if the formula contains any volatile function.
                    if (IsFormulaVolatile(formula))
                    {
                        // Record location: SheetName!CellName (e.g., "Sheet1!B2").
                        string location = $"{sheet.Name}!{cell.Name}";
                        volatileFormulaLocations.Add($"{location}: {formula}");
                    }
                }
            }

            // Output the results.
            Console.WriteLine("Volatile formulas found:");
            if (volatileFormulaLocations.Count == 0)
            {
                Console.WriteLine("None.");
            }
            else
            {
                foreach (string entry in volatileFormulaLocations)
                {
                    Console.WriteLine(entry);
                }
            }

            // Optionally, write the log to a text file next to the workbook.
            string logPath = System.IO.Path.ChangeExtension(workbookPath, ".volatile.log");
            System.IO.File.WriteAllLines(logPath, volatileFormulaLocations);
            Console.WriteLine($"Log written to: {logPath}");
        }

        // Determines whether a formula string uses any volatile function.
        private static bool IsFormulaVolatile(string formula)
        {
            if (string.IsNullOrEmpty(formula))
                return false;

            // Remove leading '=' if present for easier matching.
            string trimmed = formula.TrimStart('=').ToUpperInvariant();

            foreach (string volatileFunc in VolatileFunctions)
            {
                // Compare using case‑insensitive contains.
                if (trimmed.Contains(volatileFunc.ToUpperInvariant()))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
