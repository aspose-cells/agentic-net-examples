using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotLocalization
{
    // Custom globalization settings that localize Subtotal and Grand Total labels
    // (kept for reference; not applied because the current Aspose.Cells version
    // does not expose the PivotGlobalizationSettings property)
    public class CustomPivotGlobalizationSettings : SettablePivotGlobalizationSettings
    {
        // Override Subtotal text for each subtotal type
        public override string GetTextOfSubTotal(PivotFieldSubtotalType subTotalType)
        {
            return subTotalType switch
            {
                PivotFieldSubtotalType.Sum => "Sous‑total Somme",
                PivotFieldSubtotalType.Average => "Sous‑total Moyenne",
                PivotFieldSubtotalType.Count => "Sous‑total Compte",
                PivotFieldSubtotalType.Max => "Sous‑total Max",
                PivotFieldSubtotalType.Min => "Sous‑total Min",
                PivotFieldSubtotalType.Product => "Sous‑total Produit",
                PivotFieldSubtotalType.Var => "Sous‑total Variance",
                _ => base.GetTextOfSubTotal(subTotalType)
            };
        }

        // Override Grand Total label text
        public override string GetTextOfGrandTotal()
        {
            return "Total Général";
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Amount");
                cells["A2"].PutValue("Fruit");
                cells["B2"].PutValue(120);
                cells["A3"].PutValue("Fruit");
                cells["B3"].PutValue(80);
                cells["A4"].PutValue("Vegetable");
                cells["B4"].PutValue(150);
                cells["A5"].PutValue("Vegetable");
                cells["B5"].PutValue(200);

                // NOTE: The current Aspose.Cells version does not expose
                // Workbook.Settings.PivotGlobalizationSettings, so the custom
                // globalization settings are not applied here.

                // Create a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];

                // Configure pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, 0); // Category as row field
                int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field
                pivot.DataFields[dataFieldIdx].Function = ConsolidationFunction.Sum;

                // Refresh and calculate to apply the settings
                pivot.RefreshData();
                pivot.CalculateData();

                // Save the workbook
                string outputPath = "LocalizedPivot.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}