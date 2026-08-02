using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsPivotLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("SubCategory");
                dataSheet.Cells["C1"].PutValue("Amount");

                dataSheet.Cells["A2"].PutValue("Food");
                dataSheet.Cells["B2"].PutValue("Fruit");
                dataSheet.Cells["C2"].PutValue(120);

                dataSheet.Cells["A3"].PutValue("Food");
                dataSheet.Cells["B3"].PutValue("Vegetable");
                dataSheet.Cells["C3"].PutValue(80);

                dataSheet.Cells["A4"].PutValue("Beverage");
                dataSheet.Cells["B4"].PutValue("Tea");
                dataSheet.Cells["C4"].PutValue(50);

                dataSheet.Cells["A5"].PutValue("Beverage");
                dataSheet.Cells["B5"].PutValue("Coffee");
                dataSheet.Cells["C5"].PutValue(70);

                // Add a new worksheet for the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Create the pivot table based on the data range
                int pivotIndex = pivotSheet.PivotTables.Add("A1:C5", "A1", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category as row
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // SubCategory as column
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Amount as data

                // Create a SettablePivotGlobalizationSettings instance to customize subtotal labels
                SettablePivotGlobalizationSettings localizationSettings = new SettablePivotGlobalizationSettings();

                // Set custom texts for various subtotal types (example: localized to French)
                localizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Somme");
                localizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "Nombre");
                localizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "Moyenne");
                localizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Max, "Maximum");
                localizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Min, "Minimum");
                localizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Product, "Produit");
                // The following subtotal types are not available in the current Aspose.Cells version, so they are omitted:
                // StdDev, StdDevP, VarP

                // Set custom text for variance (available)
                localizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Var, "Variance");

                // Ensure GlobalizationSettings object exists
                if (workbook.Settings.GlobalizationSettings == null)
                    workbook.Settings.GlobalizationSettings = new GlobalizationSettings();

                // Assign the custom globalization settings to the workbook
                workbook.Settings.GlobalizationSettings.PivotSettings = localizationSettings;

                // Refresh and calculate the pivot table so that the new labels take effect
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Prepare output file path
                string outputPath = "LocalizedPivotSubtotals.xlsx";

                // Delete existing file to avoid conflicts
                if (File.Exists(outputPath))
                    File.Delete(outputPath);

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}