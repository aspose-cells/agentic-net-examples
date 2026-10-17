// Title: Locate and highlight the cell with the maximum numeric value in an Excel worksheet using Aspose.Cells Cells enumerator (C#)
// AI Prompts: Write C# code with Aspose.Cells that iterates over all cells in the first worksheet, identifies the numeric cell with the greatest value, and outputs its address. | Enhance the previous code to set a solid yellow background on the identified cell and save the workbook to a new file.
// Common Searches: aspocells c# find cell with highest number in worksheet | enumerate all cells in Excel using Aspose.Cells to get max numeric value | how to highlight the maximum numeric cell in an Excel file with Aspose.Cells C# | retrieve address of largest numeric entry in a spreadsheet using Aspose.Cells
// Tags: cells enumerator maximum value Aspose.Cells | highlight cell background Aspose.Cells C# | numeric value detection Excel Aspose.Cells | cell address extraction max number Aspose.Cells | save modified workbook Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an Excel file, uses a Cells enumerator to scan every cell in the first worksheet, tracks the highest numeric value and its cell address, prints the result, applies a yellow solid background to that cell, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Variables to keep track of the highest numeric value and its cell address
        double maxValue = double.MinValue;
        string maxCellAddress = string.Empty;

        // Enumerate all cells in the worksheet
        foreach (Cell cell in sheet.Cells)
        {
            // Consider only numeric cells
            if (cell.Type == CellValueType.IsNumeric)
            {
                double currentValue = cell.DoubleValue;
                if (currentValue > maxValue)
                {
                    maxValue = currentValue;
                    maxCellAddress = cell.Name; // e.g., "C7"
                }
            }
        }

        // Output the result
        Console.WriteLine($"Highest numeric value: {maxValue} found at cell {maxCellAddress}");

        // Optionally highlight the cell with the highest value
        if (!string.IsNullOrEmpty(maxCellAddress))
        {
            Cell maxCell = sheet.Cells[maxCellAddress];
            Style style = maxCell.GetStyle();
            style.ForegroundColor = System.Drawing.Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            maxCell.SetStyle(style);
        }

        // Save the workbook (if any modifications were made)
        workbook.Save("output.xlsx");
    }
}
