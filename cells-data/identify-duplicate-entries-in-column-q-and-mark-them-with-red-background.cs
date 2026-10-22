// Title: Highlight duplicate values in column Q of an Excel workbook with a red background using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, finds repeated values in column Q, and applies a solid red fill style to those cells. | Create a reusable C# method using Aspose.Cells that returns the duplicate cells in a specified column and colors them red.
// Common Searches: Aspose.Cells C# how to color duplicate cells in a specific column red | C# find and highlight duplicate entries in column Q of an Excel file using Aspose.Cells | Mark repeated values with red background in Excel using Aspose.Cells .NET | Detect duplicate strings in column Q and apply red fill with Aspose.Cells for .NET
// Tags: duplicate detection column Q Aspose.Cells | red background formatting Aspose.Cells | highlight repeated entries C# | Excel column duplicate marking Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;

// The program loads an input.xlsx workbook, scans column Q for duplicate string values, collects the duplicate cells, applies a solid red background style to each, and saves the result as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0]; // work with the first worksheet

        // Column Q corresponds to index 16 (zero‑based)
        int columnIndex = 16;

        // Track values that have already been seen
        HashSet<string> seenValues = new HashSet<string>();
        // Collect cells that are duplicates
        List<Cell> duplicateCells = new List<Cell>();

        // Determine the last row that contains data in the sheet
        int lastRow = sheet.Cells.MaxDataRow;

        // Scan each cell in column Q
        for (int row = 0; row <= lastRow; row++)
        {
            Cell cell = sheet.Cells[row, columnIndex];
            string cellText = cell.StringValue?.Trim();

            // Skip empty cells
            if (string.IsNullOrEmpty(cellText))
                continue;

            // If the value is already in the set, it's a duplicate
            if (!seenValues.Add(cellText))
            {
                duplicateCells.Add(cell);
            }
        }

        // Create a style with a solid red background
        Style redBackground = workbook.CreateStyle();
        redBackground.ForegroundColor = Color.Red;
        redBackground.Pattern = BackgroundType.Solid;

        // Apply the red background style to each duplicate cell
        foreach (Cell dupCell in duplicateCells)
        {
            dupCell.SetStyle(redBackground);
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
