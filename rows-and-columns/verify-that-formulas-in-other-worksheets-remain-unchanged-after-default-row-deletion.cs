// Title: Delete the first row of the first worksheet and verify that formulas in other worksheets remain unchanged using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, record every formula from all sheets, delete row 1 of the first worksheet, then compare each formula in the remaining worksheets to the recorded values to ensure none have changed, using Aspose.Cells in C#. | Using Aspose.Cells for .NET, programmatically remove the top row of the primary sheet and automatically validate that no formula cells in any other worksheet were altered after the deletion.
// Common Searches: aspnet delete first row of first sheet without affecting formulas in other worksheets Aspose.Cells | C# check if formulas change after deleting a row with Aspose.Cells | preserve cross-sheet formulas when removing rows using Aspose.Cells for .NET | how to compare Excel formulas before and after modifying a workbook in C#
// Tags: Aspose.Cells delete row keep other sheet formulas | C# capture and compare Excel formulas Aspose.Cells | verify formula stability after worksheet row deletion .NET | Excel workbook modification formula integrity Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel file, stores all formulas from every worksheet, deletes the first row of the first worksheet, then iterates through the other sheets to confirm each formula matches its original value, reports any discrepancies, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Store original formulas from all worksheets
            var originalFormulas = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
            foreach (Worksheet ws in workbook.Worksheets)
            {
                var sheetFormulas = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (Cell cell in ws.Cells)
                {
                    if (cell.IsFormula)
                    {
                        sheetFormulas[cell.Name] = cell.Formula;
                    }
                }
                originalFormulas[ws.Name] = sheetFormulas;
            }

            // Delete the first row from the first worksheet
            Worksheet targetSheet = workbook.Worksheets[0];
            targetSheet.Cells.DeleteRows(0, 1); // row index 0 = first row

            // Verify that formulas in other worksheets remain unchanged
            bool allUnchanged = true;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name == targetSheet.Name) continue;

                foreach (Cell cell in ws.Cells)
                {
                    if (cell.IsFormula)
                    {
                        string address = cell.Name;
                        if (originalFormulas.TryGetValue(ws.Name, out var sheetDict) &&
                            sheetDict.TryGetValue(address, out var before))
                        {
                            string after = cell.Formula;
                            if (!string.Equals(before, after, StringComparison.Ordinal))
                            {
                                allUnchanged = false;
                                Console.WriteLine($"Formula changed in sheet '{ws.Name}' cell {address}: before='{before}' after='{after}'");
                            }
                        }
                    }
                }
            }

            Console.WriteLine(allUnchanged ? "All formulas unchanged." : "Some formulas were altered.");

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
