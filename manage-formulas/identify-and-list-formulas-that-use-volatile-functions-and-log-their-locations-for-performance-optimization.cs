// Title: Scan an Excel workbook with Aspose.Cells (C#) to find and log cells containing volatile formulas for performance optimization
// AI Prompts: Generate C# code using Aspose.Cells that iterates through every worksheet, detects formulas with volatile functions (e.g., NOW, RAND, OFFSET), and writes their sheet and cell addresses to a text file. | Create a method in C# that returns a list of cell references where volatile Excel functions are used, leveraging Aspose.Cells' IsFormula property. | Write a .NET console application that loads a .xlsx file, searches for volatile functions in formulas, and outputs the locations to both the console and a log file.
// Common Searches: how to programmatically list cells with volatile Excel functions using Aspose.Cells in C# | Aspose.Cells C# example for detecting NOW and RAND formulas in a workbook | log addresses of volatile formulas (OFFSET, INDIRECT) in .xlsx with Aspose.Cells | performance tuning Excel files by finding volatile formulas via Aspose.Cells .NET | C# script to extract sheet name and cell address of volatile functions in Excel
// Tags: volatile-formula detection Aspose.Cells C# | log cell addresses of volatile functions .NET | scan workbook for performance‑impacting formulas Aspose.Cells | extract formula locations Excel .xlsx C# | identify volatile Excel functions using Aspose.Cells API

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates over each used cell in every worksheet, checks for formulas that contain known volatile functions (NOW, TODAY, RAND, etc.), records each occurrence as SheetName!CellAddress, prints the list to the console, and writes the locations to a log file for further performance analysis.
class VolatileFormulaFinder
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        var workbook = new Workbook("input.xlsx");

        // Collection to store locations of volatile formulas
        var volatileLocations = new List<string>();

        // Set of known volatile function names (case‑insensitive)
        var volatileFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "NOW",
            "TODAY",
            "RAND",
            "RANDBETWEEN",
            "OFFSET",
            "INDIRECT",
            "INFO",
            "CELL",
            "AREAS",
            "GETPIVOTDATA",
            "HYPERLINK"
        };

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Scan all used cells
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = sheet.Cells[row, col];

                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Check if the formula contains any volatile function
                        foreach (string func in volatileFunctions)
                        {
                            // Look for the function name followed by '(' to avoid false matches
                            if (formula.IndexOf(func + "(", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Record the location as SheetName!CellAddress (e.g., Sheet1!A1)
                                volatileLocations.Add($"{sheet.Name}!{cell.Name}");
                                break; // No need to check other functions for this cell
                            }
                        }
                    }
                }
            }
        }

        // Output the results to the console
        Console.WriteLine("Volatile formulas found at:");
        foreach (string location in volatileLocations)
        {
            Console.WriteLine(location);
        }

        // Optionally, write the locations to a log file for further analysis
        System.IO.File.WriteAllLines("VolatileFormulasLog.txt", volatileLocations);
    }
}
