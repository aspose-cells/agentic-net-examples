using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCustomGlobalization
{
    // Subclass of GlobalizationSettings that provides localized total names.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override GetTotalName to return custom labels based on the consolidation function.
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            return functionType switch
            {
                ConsolidationFunction.Sum => "Localized Sum",
                ConsolidationFunction.Average => "Localized Average",
                ConsolidationFunction.Count => "Localized Count",
                ConsolidationFunction.Max => "Localized Max",
                ConsolidationFunction.Min => "Localized Min",
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
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for a pivot table.
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("A");
            cells["B4"].PutValue(30);
            cells["A5"].PutValue("B");
            cells["B5"].PutValue(40);

            // Apply the custom globalization settings to the workbook.
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Create a pivot table to demonstrate the custom total names.
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SamplePivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, 1); // Value
            // Set the function to Sum; the overridden GetTotalName will be used for the total label.
            pivot.DataFields[dataFieldIdx].Function = ConsolidationFunction.Sum;

            // Refresh and calculate the pivot table.
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook.
            workbook.Save("CustomGlobalizationSettingsDemo.xlsx");
        }
    }
}