// Title: Identify and list Excel cells that contain volatile functions using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, scans every worksheet, and returns a list of addresses for cells whose formulas include volatile functions such as NOW, TODAY, RAND, OFFSET, etc. | Create a reusable C# method that takes a file path and outputs a dictionary mapping each sheet name to the cell addresses of formulas that use any volatile Excel function. | Extend the sample to also report which specific volatile function was detected in each identified cell together with its address.
// Common Searches: aspocells c# find cells with volatile formulas in an Excel workbook | list addresses of cells that use NOW or RAND functions using Aspose.Cells .NET | how to detect volatile Excel functions programmatically with Aspose.Cells | iterate used range and check for volatile functions in C# Aspose.Cells example
// Tags: Aspose.Cells workbook scanning for volatile functions | C# extract cell addresses of volatile formulas | used range traversal Aspose.Cells .NET | Excel formula analysis with Aspose.Cells | detect Excel volatile calculations programmatically

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, iterates each worksheet's used range, checks every formula cell for a set of known volatile functions (NOW, TODAY, RAND, OFFSET, etc.), records the sheet name and cell address of matches, and prints the collected list.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // List to store addresses of cells that use volatile functions
            List<string> volatileCellAddresses = new List<string>();

            // Common volatile functions in Excel (uppercase for case‑insensitive comparison)
            string[] volatileFunctions = new string[]
            {
                "NOW()", "TODAY()", "RAND()", "RANDBETWEEN(", "OFFSET(", "INDIRECT(", "INFO(", "CELL(", "AREAS(", "GETPIVOTDATA("
            };

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use the used range to limit the iteration to populated cells
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // If the worksheet is empty, skip it
                if (usedRange == null)
                    continue;

                // Examine each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Check if the cell contains a formula
                    if (cell.IsFormula)
                    {
                        // Get the formula text in a uniform case
                        string formula = cell.Formula.ToUpperInvariant();

                        // Check if the formula contains any volatile function
                        foreach (string vf in volatileFunctions)
                        {
                            if (formula.Contains(vf))
                            {
                                // Record the address with sheet name (e.g., Sheet1!A1)
                                volatileCellAddresses.Add($"{sheet.Name}!{cell.Name}");
                                break; // No need to check other functions for this cell
                            }
                        }
                    }
                }
            }

            // Output the collected addresses
            Console.WriteLine("Cells containing volatile formulas:");
            foreach (string address in volatileCellAddresses)
            {
                Console.WriteLine(address);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
