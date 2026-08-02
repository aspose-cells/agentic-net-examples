using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCustomGlobalization
{
    // Custom globalization settings that provides a localized label for totals.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override GetTotalName to return a custom label based on the consolidation function.
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            // Example: return a localized label for the Sum function; otherwise fallback to base implementation.
            return functionType switch
            {
                ConsolidationFunction.Sum => "Localized Subtotal (Sum)",
                ConsolidationFunction.Count => "Localized Subtotal (Count)",
                ConsolidationFunction.Average => "Localized Subtotal (Average)",
                _ => base.GetTotalName(functionType)
            };
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Assign the custom globalization settings to the workbook.
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Populate sample data for a pivot table.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("Food");
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("Food");
            cells["B3"].PutValue(80);
            cells["A4"].PutValue("Travel");
            cells["B4"].PutValue(200);
            cells["A5"].PutValue("Travel");
            cells["B5"].PutValue(150);

            // Create a pivot table to trigger the custom total name.
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SamplePivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0); // Category
            int dataFieldIdx = pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Amount
            PivotField dataField = pivotTable.DataFields[dataFieldIdx];
            dataField.Function = ConsolidationFunction.Sum; // Use Sum to see custom label

            // Refresh and calculate the pivot table.
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook.
            workbook.Save("CustomGlobalizationSettingsDemo.xlsx");
        }
    }
}