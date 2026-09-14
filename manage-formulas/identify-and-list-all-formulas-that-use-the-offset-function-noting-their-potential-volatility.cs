// Title: Find and list all Excel formulas that use the OFFSET function and mark them as volatile with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to scan every worksheet in a workbook, detect formulas containing the OFFSET function, and output each cell’s address, formula, and a volatility label. | Extend the sample to count how many OFFSET formulas appear on each worksheet and display a summary after listing the individual cells. | Create a reusable C# method that accepts any function name (e.g., OFFSET, INDIRECT) and returns a collection of cell addresses and formulas that include that function, leveraging Aspose.Cells.
// Common Searches: aspnet find OFFSET formulas in an Excel workbook using Aspose.Cells | c# list volatile Excel functions with Aspose.Cells library | how to detect OFFSET function in cell formulas with Aspose.Cells .NET | enumerate formulas that contain OFFSET in all sheets using Aspose.Cells | retrieve cell address and formula for volatile functions in Excel via C#
// Tags: scan worksheets for OFFSET formulas Aspose.Cells | detect volatile functions in Excel .NET | list cell formulas containing specific function C# | Aspose.Cells formula analysis offset | enumerate volatile Excel functions programmatically

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook, iterates through each worksheet and cell, checks each formula for the OFFSET function (case‑insensitive), records the sheet name, cell address, formula and notes that OFFSET is volatile, then prints the collected information.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // List to hold information about formulas using OFFSET
        List<string> offsetFormulas = new List<string>();

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the cells collection of the current worksheet
            Cells cells = sheet.Cells;

            // Iterate through all used cells
            foreach (Cell cell in cells)
            {
                // Check if the cell contains a formula
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Determine if the formula uses the OFFSET function (case‑insensitive)
                    if (cell.Formula.IndexOf("OFFSET(", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Build a description including sheet name, cell address, formula and volatility note
                        string info = $"Sheet: {sheet.Name}, Cell: {cell.Name}, Formula: {cell.Formula}, Volatility: Volatile (OFFSET is a volatile function)";
                        offsetFormulas.Add(info);
                    }
                }
            }
        }

        // Output the results
        Console.WriteLine("Formulas that use the OFFSET function:");
        if (offsetFormulas.Count == 0)
        {
            Console.WriteLine("None found.");
        }
        else
        {
            foreach (string entry in offsetFormulas)
            {
                Console.WriteLine(entry);
            }
        }
    }
}
