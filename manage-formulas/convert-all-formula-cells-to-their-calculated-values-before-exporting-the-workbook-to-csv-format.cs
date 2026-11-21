// Title: Convert all formula cells to static values before exporting a workbook to CSV with Aspose.Cells for .NET
// Common Searches: how to export Excel to CSV with only values using Aspose.Cells C# | Aspose.Cells replace formula cells with their calculated results before CSV save | C# code to evaluate all formulas and remove them prior to CSV conversion in Aspose.Cells | convert workbook to CSV without formulas Aspose.Cells .NET example | calculate formulas then save as CSV using Aspose.Cells library
// Tags: calculate all formulas Aspose.Cells | replace formula cells with values Aspose.Cells | CSV export without formulas Aspose.Cells | Workbook.SaveFormat.CSV static values | cell.IsFormula value replacement Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook, forces calculation of every formula, iterates through each worksheet and cell to replace any formula with its evaluated value, and then saves the workbook as a CSV file that contains only static values.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure all formulas are calculated
        workbook.CalculateFormula();

        // Replace each formula with its calculated value
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.IsFormula)
                {
                    // Put the already calculated value back into the cell, removing the formula
                    cell.PutValue(cell.Value);
                }
            }
        }

        // Export the workbook (now containing only values) to CSV format
        workbook.Save("output.csv", SaveFormat.CSV);
    }
}
