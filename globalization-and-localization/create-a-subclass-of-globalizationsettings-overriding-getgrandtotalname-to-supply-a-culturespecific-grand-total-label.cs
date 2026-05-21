using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Custom globalization settings that returns a culture‑specific grand total label
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override the method that provides the grand total name for a given consolidation function
        public override string GetGrandTotalName(ConsolidationFunction functionType)
        {
            // Example: use different labels for Sum and other functions based on current culture
            // (In a real scenario you could look up resources per culture.)
            if (functionType == ConsolidationFunction.Sum)
            {
                // English label
                return "Grand Total (Sum)";
            }
            else if (functionType == ConsolidationFunction.Average)
            {
                // French label
                return "Total Général (Moyenne)";
            }
            else
            {
                // Fallback to the base implementation for all other functions
                return base.GetGrandTotalName(functionType);
            }
        }
    }

    public class GlobalizationSettingsDemo
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Food");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Drink");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Drink");
            sheet.Cells["B5"].PutValue(200);

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Create a pivot table to demonstrate the custom grand total label
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, 1); // Amount
            pivot.DataFields[dataFieldIdx].Function = ConsolidationFunction.Sum;

            // Refresh and calculate the pivot table so that labels are generated
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("CustomGrandTotalDemo.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            GlobalizationSettingsDemo.Run();
        }
    }
}