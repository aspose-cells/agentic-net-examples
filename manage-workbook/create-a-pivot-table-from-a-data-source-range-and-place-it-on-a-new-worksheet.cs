using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    public class CreatePivotTable
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet and add sample data
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Populate sample data
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Product");
                sourceSheet.Cells["C1"].PutValue("Sales");

                sourceSheet.Cells["A2"].PutValue("Fruit");
                sourceSheet.Cells["B2"].PutValue("Apple");
                sourceSheet.Cells["C2"].PutValue(1200);

                sourceSheet.Cells["A3"].PutValue("Fruit");
                sourceSheet.Cells["B3"].PutValue("Banana");
                sourceSheet.Cells["C3"].PutValue(850);

                sourceSheet.Cells["A4"].PutValue("Vegetable");
                sourceSheet.Cells["B4"].PutValue("Carrot");
                sourceSheet.Cells["C4"].PutValue(560);

                sourceSheet.Cells["A5"].PutValue("Vegetable");
                sourceSheet.Cells["B5"].PutValue("Broccoli");
                sourceSheet.Cells["C5"].PutValue(430);

                // Determine the source data range (including headers)
                Aspose.Cells.Range sourceRange = sourceSheet.Cells.MaxDisplayRange;
                string sourceData = $"=SourceData!{sourceRange.Address}";

                // Add a new worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Add the pivot table to the new worksheet
                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceData, "A1", "SalesPivot");

                // Configure the pivot table fields
                PivotTable pivotTable = pivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh the pivot table to calculate data
                pivotSheet.RefreshPivotTables();

                // Save the workbook (overwrite if exists)
                string outputPath = "PivotTableResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}