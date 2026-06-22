using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    // Custom globalization settings for pivot tables.
    // Overrides the text for subtotals and the grand total label.
    public class CustomPivotGlobalizationSettings : PivotGlobalizationSettings
    {
        // Override subtotal text based on the subtotal type.
        public override string GetTextOfSubTotal(PivotFieldSubtotalType subTotalType)
        {
            switch (subTotalType)
            {
                case PivotFieldSubtotalType.Sum:
                    return "Custom Sum Subtotal";
                case PivotFieldSubtotalType.Average:
                    return "Custom Average Subtotal";
                case PivotFieldSubtotalType.Count:
                    return "Custom Count Subtotal";
                case PivotFieldSubtotalType.Max:
                    return "Custom Max Subtotal";
                case PivotFieldSubtotalType.Min:
                    return "Custom Min Subtotal";
                default:
                    // Fallback to default implementation for any other types.
                    return base.GetTextOfSubTotal(subTotalType);
            }
        }

        // Override the grand total label.
        public override string GetTextOfGrandTotal()
        {
            return "Custom Grand Total";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table.
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("A");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("B");
            cells["B4"].PutValue(30);
            cells["A5"].PutValue("B");
            cells["B5"].PutValue(40);

            // Create a pivot table based on the data range.
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

            // Apply custom globalization settings to the workbook.
            workbook.Settings.GlobalizationSettings.PivotSettings = new CustomPivotGlobalizationSettings();

            // Refresh and calculate the pivot table so that the custom labels are used.
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook.
            workbook.Save("LocalizedPivotTable.xlsx");
        }
    }
}