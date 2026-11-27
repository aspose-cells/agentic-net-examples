// Title: Round numeric cells to two decimal places and export an Excel workbook to CSV using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells round cell values to two decimals before saving as CSV | How to export Excel to CSV with two‑decimal precision using Aspose.Cells | Round financial numbers in a workbook and generate CSV with Aspose.Cells .NET | Aspose.Cells CSV export with numeric rounding for financial reporting | Convert .xlsx to .csv while rounding numbers to 2 decimal places in C#
// Tags: numeric rounding Aspose.Cells CSV | Excel to CSV conversion with two‑decimal precision | C# financial data rounding Aspose.Cells | midpoint rounding away from zero in .NET | process all worksheets Aspose.Cells export

using Aspose.Cells;
using System;

// The example loads an Excel workbook, iterates through every worksheet and cell, rounds any numeric value to two decimal places using Math.Round with MidpointRounding.AwayFromZero, updates the cell, and then saves the workbook (or a selected sheet) as a CSV file.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range of rows and columns
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Loop through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell contains a numeric value
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        // Round the numeric value to two decimal places
                        double original = cell.DoubleValue;
                        double rounded = Math.Round(original, 2, MidpointRounding.AwayFromZero);
                        cell.PutValue(rounded);
                    }
                }
            }
        }

        // Export the workbook (or a specific worksheet) to CSV
        // This example saves the first worksheet; adjust as needed
        workbook.Save("output.csv", SaveFormat.CSV);
    }
}
