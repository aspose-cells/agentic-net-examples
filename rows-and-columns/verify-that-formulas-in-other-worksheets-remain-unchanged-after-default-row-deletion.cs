// Title: Check that formulas in other worksheets stay unchanged after deleting the first row of the default sheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells verify formulas are unchanged after deleting a row in C# | C# compare worksheet formulas before and after row deletion with Aspose.Cells | detect modified formulas across multiple sheets after removing first row using Aspose.Cells | how to ensure Excel formulas remain intact in non-default sheets when deleting rows programmatically
// Tags: Aspose.Cells row deletion formula preservation | compare worksheet formulas Aspose.Cells C# | track Excel formula changes across sheets | validate formula integrity after modifying workbook | store and verify cell formulas Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example loads an Excel workbook, captures all formulas from every worksheet except the first, deletes the first row of the default worksheet, then iterates through the remaining sheets to confirm each stored formula matches the current one, reporting any changes, missing, or new formulas before saving the updated workbook.
    class VerifyFormulasAfterRowDeletion
    {
        static void Main(string[] args)
        {
            // Default file paths; can be overridden by command‑line arguments
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            if (args.Length >= 1) inputPath = args[0];
            if (args.Length >= 2) outputPath = args[1];

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook (rule: load-workbook)
                Workbook workbook = new Workbook(inputPath);

                // Store formulas from all worksheets except the default (first) worksheet
                var originalFormulas = new Dictionary<string, Dictionary<string, string>>();

                for (int wsIndex = 1; wsIndex < workbook.Worksheets.Count; wsIndex++)
                {
                    Worksheet ws = workbook.Worksheets[wsIndex];
                    var sheetFormulas = new Dictionary<string, string>();

                    // Iterate through all used cells
                    foreach (Cell cell in ws.Cells)
                    {
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            // Use cell name (e.g., "A1") as key
                            sheetFormulas[cell.Name] = cell.Formula;
                        }
                    }

                    originalFormulas[ws.Name] = sheetFormulas;
                }

                // Delete the default row (first row) from the default worksheet (rule: default-row-deletion)
                Worksheet defaultSheet = workbook.Worksheets[0];
                defaultSheet.Cells.DeleteRows(0, 1); // deletes row 1 (zero‑based index)

                // Verify that formulas in other worksheets remain unchanged
                bool allUnchanged = true;

                for (int wsIndex = 1; wsIndex < workbook.Worksheets.Count; wsIndex++)
                {
                    Worksheet ws = workbook.Worksheets[wsIndex];
                    var sheetFormulas = originalFormulas[ws.Name];

                    foreach (Cell cell in ws.Cells)
                    {
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            if (sheetFormulas.TryGetValue(cell.Name, out string originalFormula))
                            {
                                if (!string.Equals(cell.Formula, originalFormula, StringComparison.Ordinal))
                                {
                                    Console.WriteLine($"Formula changed in sheet '{ws.Name}' cell {cell.Name}:");
                                    Console.WriteLine($"  Original: {originalFormula}");
                                    Console.WriteLine($"  Current : {cell.Formula}");
                                    allUnchanged = false;
                                }
                            }
                            else
                            {
                                // New formula appeared after deletion
                                Console.WriteLine($"New formula detected in sheet '{ws.Name}' cell {cell.Name}: {cell.Formula}");
                                allUnchanged = false;
                            }
                        }
                    }

                    // Check for formulas that disappeared
                    foreach (var kvp in sheetFormulas)
                    {
                        if (string.IsNullOrEmpty(ws.Cells[kvp.Key].Formula))
                        {
                            Console.WriteLine($"Formula missing in sheet '{ws.Name}' cell {kvp.Key}: was {kvp.Value}");
                            allUnchanged = false;
                        }
                    }
                }

                Console.WriteLine(allUnchanged
                    ? "All formulas in other worksheets remain unchanged after row deletion."
                    : "Some formulas were altered after row deletion.");

                // Save the workbook (rule: save-workbook)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
