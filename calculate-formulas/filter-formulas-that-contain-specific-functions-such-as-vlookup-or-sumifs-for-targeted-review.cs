// Title: C# Example: Filter Cells Containing VLOOKUP or SUMIFS Formulas with Aspose.Cells
// Description: Loads an Excel workbook, scans the first worksheet for formulas that include VLOOKUP or SUMIFS using Aspose.Cells FindOptions (OnlyFormulas + Contains), gathers each unique cell, prints its address and formula, and saves the workbook unchanged. Ideal for .NET developers needing to audit or extract specific functions from Excel files.
// Keywords: Aspose.Cells | C# | .NET | find formulas | VLOOKUP | SUMIFS | filter cells by function | Excel workbook analysis | FindOptions OnlyFormulas | code example | GitHub
// Common Searches: Aspose.Cells find VLOOKUP formulas C# | search for SUMIFS cells using Aspose.Cells .NET | list Excel cells that contain specific functions | avoid duplicate matches when searching multiple formulas | C# example to audit lookup functions in a workbook
// Developer Intent: Locate and list every cell whose formula includes VLOOKUP or SUMIFS.
// Use Cases: Create an audit report of all lookup and conditional‑sum formulas in a spreadsheet. | Validate that prohibited functions are not present before distributing an Excel file. | Extract formula strings for bulk refactoring or migration to newer functions.
// AI Prompts: Generate C# code with Aspose.Cells that finds all INDEX function formulas and outputs their addresses. | Show how to replace each VLOOKUP formula with an XLOOKUP equivalent using Aspose.Cells. | Provide a snippet that logs matched cells to a CSV file while keeping the original workbook unchanged.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace FormulaFilterDemo
{
    // Loads an Excel workbook, scans the first worksheet for formulas that include VLOOKUP or SUMIFS using Aspose.Cells FindOptions (OnlyFormulas + Contains), gathers each unique cell, prints its address and formula, and saves the workbook unchanged. Ideal for .NET developers needing to audit or extract specific functions from Excel files.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input workbook path
                string inputPath = "input.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Work with the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Functions we want to locate in formulas
                string[] targetFunctions = { "VLOOKUP", "SUMIFS" };

                // Store cells that contain any of the target functions
                List<Cell> matchedCells = new List<Cell>();

                foreach (string func in targetFunctions)
                {
                    // Set find options to search only in formulas and allow partial matches
                    FindOptions options = new FindOptions
                    {
                        LookInType = LookInType.OnlyFormulas,
                        LookAtType = LookAtType.Contains
                    };

                    // Find the first occurrence of the function
                    Cell found = worksheet.Cells.Find(func, null, options);

                    // Continue searching until no more matches are found
                    while (found != null)
                    {
                        // Avoid duplicate entries when a cell matches multiple functions
                        if (!matchedCells.Contains(found))
                            matchedCells.Add(found);

                        // Find the next occurrence starting after the current cell
                        found = worksheet.Cells.Find(func, found, options);
                    }
                }

                // Output the addresses and formulas of the matched cells
                Console.WriteLine("Cells containing VLOOKUP or SUMIFS formulas:");
                foreach (Cell cell in matchedCells)
                {
                    Console.WriteLine($"- {cell.Name}: {cell.Formula}");
                }

                // Save the workbook (unchanged in this demo)
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
