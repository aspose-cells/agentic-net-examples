// Title: Check and report cells exceeding 255 characters with Aspose.Cells when Excel compatibility mode is on (C#)
// Common Searches: C# Aspose.Cells iterate used range to find strings longer than 255 characters | how to enforce Excel 97-2003 compatibility limits with Aspose.Cells in .NET | detect cells that violate 255‑character limit after enabling CheckCompatibility in Aspose.Cells | list cell addresses with text length over 255 using Aspose.Cells C# | validate workbook for legacy Excel string length restrictions with Aspose.Cells
// Tags: Aspose.Cells compatibility mode validation | cell string length validation Aspose.Cells | used range iteration Aspose.Cells C# | legacy Excel format string limit Aspose | report cells exceeding 255 characters Aspose

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads a workbook, enables CheckCompatibility to apply Excel 97‑2003 limits, scans each worksheet's used range for string cells longer than 255 characters, records the sheet name and cell address of each violation, prints a report, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Enable compatibility checking (Excel 97-2003 limits)
        workbook.Settings.CheckCompatibility = true;

        // Collect cells that violate the 255‑character limit
        List<string> violations = new List<string>();

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate only over the used range to improve performance
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Check only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue;
                        if (!string.IsNullOrEmpty(text) && text.Length > 255)
                        {
                            // Record the violation with sheet name and cell address
                            violations.Add(
                                $"Sheet \"{sheet.Name}\", Cell {cell.Name}: length {text.Length} exceeds 255 characters.");
                        }
                    }
                }
            }
        }

        // Report the result
        if (violations.Count > 0)
        {
            Console.WriteLine("Compatibility check failed. Cells exceeding 255 characters:");
            foreach (string v in violations)
                Console.WriteLine(v);
        }
        else
        {
            Console.WriteLine("All cell contents are within the 255‑character limit.");
        }

        // Save the workbook if further processing is required
        workbook.Save("output.xlsx");
    }
}
