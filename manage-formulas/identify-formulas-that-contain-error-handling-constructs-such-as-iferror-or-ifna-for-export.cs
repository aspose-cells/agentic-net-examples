// Title: Identify and export Excel formulas that use IFERROR or IFNA with Aspose.Cells for .NET (C#)
// Common Searches: how to find and export IFERROR formulas from an .xlsx using Aspose.Cells C# | C# code to list cells containing IFNA in a workbook with Aspose.Cells | extract error‑handling formulas from Excel and save to a text file in .NET | Aspose.Cells iterate through worksheets and capture formulas with IFERROR | export formulas that contain IFERROR or IFNA to a .txt file using C#
// Tags: Aspose.Cells error‑handling formula extraction | C# export workbook formulas to txt | identify IFNA functions in Excel cells | iterate worksheets and cells with Aspose.Cells | filter formulas containing error functions .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, walks through every worksheet and cell, detects formulas that contain the IFERROR or IFNA functions (case‑insensitive), records each match with its sheet name and address, and writes the list to a plain‑text file.
class IdentifyErrorHandlingFormulas
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Path where the list of formulas will be exported
        string outputPath = "ErrorHandlingFormulas.txt";

        // Load the workbook (Aspose.Cells handles all supported formats)
        Workbook workbook = new Workbook(inputPath);

        // Collection to store formulas that contain IFERROR or IFNA
        List<string> errorFormulas = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate through each cell that has a formula
            foreach (Cell cell in cells)
            {
                // Ensure the cell actually contains a formula
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    string formula = cell.Formula;

                    // Check for error‑handling functions (case‑insensitive)
                    if (formula.IndexOf("IFERROR", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        formula.IndexOf("IFNA", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Record the sheet name, cell address, and the formula itself
                        errorFormulas.Add($"{sheet.Name}!{cell.Name}: {formula}");
                    }
                }
            }
        }

        // Export the collected formulas to a text file (one per line)
        File.WriteAllLines(outputPath, errorFormulas);

        // Optional: inform the user
        Console.WriteLine($"Exported {errorFormulas.Count} formulas containing IFERROR/IFNA to '{outputPath}'.");
    }
}
