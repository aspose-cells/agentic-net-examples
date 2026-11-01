// Title: Detect formulas that reference hidden worksheets in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# code to list cells that contain formulas pointing to hidden sheets using Aspose.Cells | how to audit Excel workbook for hidden worksheet references in formulas with Aspose.Cells .NET | find hidden sheet formula links in a .xlsx file using Aspose.Cells C# | scan all formulas for hidden worksheet references in Aspose.Cells for .NET | identify cells referencing hidden worksheets for security review in C#
// Tags: detect hidden worksheet references Aspose.Cells | scan workbook formulas for hidden sheets .NET | audit Excel formulas hidden sheet links C# | enumerate cells with hidden sheet references Aspose.Cells | security review hidden worksheet formulas Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, gathers the names of hidden worksheets, iterates through every cell in each sheet, checks each formula for references to any hidden sheet (case‑insensitive), and outputs the addresses of cells that contain such references for security auditing.
class DetectHiddenSheetFormulaReferences
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Ensure the file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect names of hidden worksheets (case‑insensitive)
            HashSet<string> hiddenSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (!ws.IsVisible) // IsVisible == false means hidden
                {
                    hiddenSheetNames.Add(ws.Name);
                }
            }

            // List to store cells that reference hidden worksheets
            List<string> offendingCells = new List<string>();

            // Scan all cells in all worksheets for formulas referencing hidden sheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells cells = ws.Cells;
                foreach (Cell cell in cells)
                {
                    // Process only cells that contain a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        string formula = cell.Formula;

                        // Check if the formula contains a reference to any hidden sheet
                        foreach (string hiddenSheet in hiddenSheetNames)
                        {
                            // Look for the pattern "SheetName!" (case‑insensitive)
                            if (formula.IndexOf(hiddenSheet + "!", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                offendingCells.Add($"{ws.Name}!{cell.Name}");
                                break; // No need to check other hidden sheets for this cell
                            }
                        }
                    }
                }
            }

            // Output the results
            if (offendingCells.Count > 0)
            {
                Console.WriteLine("Formulas referencing hidden worksheets were found in the following cells:");
                foreach (string address in offendingCells)
                {
                    Console.WriteLine(address);
                }
            }
            else
            {
                Console.WriteLine("No formulas referencing hidden worksheets were detected.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
