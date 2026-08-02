using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCustomGrandTotal
{
    // Subclass of GlobalizationSettings that provides a culture‑specific grand total label
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override GetGrandTotalName to return a custom label based on the consolidation function
        public override string GetGrandTotalName(ConsolidationFunction functionType)
        {
            // Example: French culture labels
            return functionType switch
            {
                ConsolidationFunction.Sum => "Total Général (Somme)",
                ConsolidationFunction.Average => "Total Général (Moyenne)",
                ConsolidationFunction.Count => "Total Général (Compte)",
                _ => base.GetGrandTotalName(functionType)
            };
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue(200);

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Create a pivot table to demonstrate the custom grand total label
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);          // Category
            int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, 1); // Amount
            pivot.DataFields[dataFieldIdx].Function = ConsolidationFunction.Sum;

            // Refresh and calculate the pivot table so that labels are generated
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook (save rule)
            workbook.Save("CustomGrandTotalDemo.xlsx");
        }
    }
}