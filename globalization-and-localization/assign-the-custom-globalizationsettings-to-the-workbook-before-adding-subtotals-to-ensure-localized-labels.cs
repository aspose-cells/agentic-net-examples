using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Amount");
        cells["A2"].PutValue("A");
        cells["B2"].PutValue(10);
        cells["A3"].PutValue("A");
        cells["B3"].PutValue(20);
        cells["A4"].PutValue("B");
        cells["B4"].PutValue(30);
        cells["A5"].PutValue("B");
        cells["B5"].PutValue(40);

        // Assign custom globalization settings before adding subtotals
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Define the range for subtotal (rows 0‑4, columns 0‑1)
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 1);

        // Add subtotal: group by column 0 (Category) and calculate Sum on column 1 (Amount)
        // Parameters: area, columnIndexToGroup, function, columnsToSubtotal, replace, pageBreaks, summaryBelowData
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

        // Save the workbook
        workbook.Save("CustomGlobalizationSubtotal.xlsx");
    }

    // Custom globalization settings to provide localized total label
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            // Return a custom label for the Sum total; fall back to default for others
            return functionType == ConsolidationFunction.Sum
                ? "Custom Sum Total"
                : base.GetTotalName(functionType);
        }
    }
}