// Title: Collect cells with #DIV/0! errors using Aspose.Cells for .NET
// Description: Loads an Excel workbook, calculates all formulas, scans each worksheet, detects formula cells whose result is the #DIV/0! error, records their addresses, prints them, and saves the workbook.
// Keywords: Aspose.Cells #DIV/0! detection | C# find division by zero error cells | Excel formula error extraction Aspose | list error cells Aspose.Cells | retrieve error values .NET
// Common Searches: Aspose.Cells find #DIV/0! cells | C# get Excel cells with division by zero error | list formula error locations using Aspose | detect #DIV/0! after workbook.CalculateFormula | how to enumerate error cells in Aspose.Cells
// Developer Intent: Locate every formula cell that evaluates to #DIV/0! and obtain its address for further processing.
// Use Cases: Generate an error report before publishing a workbook | Automatically replace #DIV/0! with a default value | Log problematic cell locations for debugging complex spreadsheets | Highlight error cells in a UI for end‑user review
// AI Prompts: Write C# code with Aspose.Cells that scans a workbook, identifies all cells returning #DIV/0! after calculation, and returns their addresses. | Provide an Aspose.Cells snippet that iterates through each worksheet, checks for error values, and collects cells with the #DIV/0! error into a list. | Explain how to modify the collected #DIV/0! cells to set a custom placeholder text using Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook, calculates all formulas, scans each worksheet, detects formula cells whose result is the #DIV/0! error, records their addresses, prints them, and saves the workbook.
class IdentifyDiv0Errors
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // List to store cells that contain a #DIV/0! error
        List<Cell> div0ErrorCells = new List<Cell>();

        // Scan every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate through all cells in the used range
            foreach (Cell cell in cells)
            {
                // Process only formula cells
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // After calculation, check if the cell result is an error and specifically #DIV/0!
                    if (cell.IsErrorValue && cell.StringValue == "#DIV/0!")
                    {
                        div0ErrorCells.Add(cell);
                    }
                }
            }
        }

        // Output the addresses of cells with #DIV/0! errors
        Console.WriteLine("Cells containing #DIV/0! error:");
        foreach (Cell errCell in div0ErrorCells)
        {
            Console.WriteLine($"{errCell.Name} (Worksheet: {errCell.Worksheet.Name})");
        }

        // Save the workbook if further processing is needed
        workbook.Save("output.xlsx");
    }
}
