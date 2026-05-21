using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    public class LocalizedPivotSubtotalsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Amount");
                dataSheet.Cells["A2"].PutValue("Food");
                dataSheet.Cells["B2"].PutValue(120);
                dataSheet.Cells["A3"].PutValue("Food");
                dataSheet.Cells["B3"].PutValue(80);
                dataSheet.Cells["A4"].PutValue("Drink");
                dataSheet.Cells["B4"].PutValue(150);
                dataSheet.Cells["A5"].PutValue("Drink");
                dataSheet.Cells["B5"].PutValue(70);

                // Add a pivot table based on the data range
                int pivotIndex = dataSheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
                PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

                // Customize subtotal texts
                SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Σ Total");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "Count Total");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "Avg Total");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Max, "Maximum");
                pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Min, "Minimum");

                // Apply settings to workbook
                workbook.Settings.GlobalizationSettings.PivotSettings = pivotSettings;

                // Refresh and calculate pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save workbook
                string outputPath = "LocalizedPivotSubtotals.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point
        public static void Main(string[] args)
        {
            Run();
        }
    }
}