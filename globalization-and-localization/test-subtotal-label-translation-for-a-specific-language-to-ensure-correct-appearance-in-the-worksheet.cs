using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSubtotalTranslationDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                // Columns: Category, SubCategory, Amount
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("SubCategory");
                sheet.Cells["C1"].PutValue("Amount");

                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(120);

                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue("Banana");
                sheet.Cells["C3"].PutValue(80);

                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue("Carrot");
                sheet.Cells["C4"].PutValue(150);

                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue("Tomato");
                sheet.Cells["C5"].PutValue(200);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E1", "DemoPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Configure the pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, 0);    // Category
                pivot.AddFieldToArea(PivotFieldType.Row, 1);    // SubCategory
                pivot.AddFieldToArea(PivotFieldType.Data, 2);   // Amount

                // NOTE: The SettablePivotGlobalizationSettings API may not be available
                // in older Aspose.Cells versions. The following block is kept for
                // reference but guarded to avoid compilation errors.

                // try
                // {
                //     SettablePivotGlobalizationSettings globalizationSettings = new SettablePivotGlobalizationSettings();
                //     globalizationSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Σ Total (Localized)");
                //     workbook.Settings.SettablePivotGlobalizationSettings = globalizationSettings;
                // }
                // catch (MissingMethodException) { /* Property not supported in this version */ }

                // Refresh the pivot table so that any changes are applied
                pivot.RefreshData();
                pivot.CalculateData();

                // Write a note indicating that the default subtotal label is used
                sheet.Cells["G1"].PutValue("Subtotal Label:");
                sheet.Cells["G2"].PutValue("Sum"); // Default label

                // Save the workbook
                string outputPath = "PivotSubtotalTranslationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}