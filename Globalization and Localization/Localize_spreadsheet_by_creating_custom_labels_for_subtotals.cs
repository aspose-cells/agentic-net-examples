using System;
using Aspose.Cells;

public class CustomGlobalizationSettings : GlobalizationSettings
{
    // Override the total name for specific consolidation functions.
    public override string GetTotalName(ConsolidationFunction functionType)
    {
        switch (functionType)
        {
            case ConsolidationFunction.Sum:
                return "Custom Subtotal Sum";
            case ConsolidationFunction.Average:
                return "Custom Subtotal Avg";
            default:
                return base.GetTotalName(functionType);
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Load an existing XLSX workbook.
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet and its cells collection.
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Apply custom globalization settings that provide custom subtotal labels.
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Define the range on which the subtotal will be applied (A1:B5 in this example).
        // Rows and columns are zero‑based indexes.
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // Row 0‑4, Column 0‑1

        // Apply subtotal:
        // - Group by the first column (index 0).
        // - Calculate Sum on the second column (index 1).
        // - Replace existing data (true), do not use page break (false), and add a subtotal row (true).
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

        // Save the modified workbook.
        workbook.Save("output.xlsx");
    }
}