using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace PivotTableExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ----- Original sheet with source data -----
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Populate sample source data
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Product");
                sourceSheet.Cells["C1"].PutValue("Sales");

                sourceSheet.Cells["A2"].PutValue("Fruit");
                sourceSheet.Cells["B2"].PutValue("Apple");
                sourceSheet.Cells["C2"].PutValue(1200);

                sourceSheet.Cells["A3"].PutValue("Fruit");
                sourceSheet.Cells["B3"].PutValue("Orange");
                sourceSheet.Cells["C3"].PutValue(850);

                sourceSheet.Cells["A4"].PutValue("Vegetable");
                sourceSheet.Cells["B4"].PutValue("Carrot");
                sourceSheet.Cells["C4"].PutValue(560);

                sourceSheet.Cells["A5"].PutValue("Vegetable");
                sourceSheet.Cells["B5"].PutValue("Broccoli");
                sourceSheet.Cells["C5"].PutValue(430);

                // ----- New worksheet for the pivot table -----
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Determine the source data range (including headers)
                AsposeRange sourceRange = sourceSheet.Cells.MaxDisplayRange;
                string sourceData = $"=SourceData!{sourceRange.Address}";

                // Add a pivot table to the new sheet (destination cell A1, name "SalesPivot")
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Optional: apply a style for better appearance
                pivotTable.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

                // Refresh the pivot table to calculate data
                pivotSheet.RefreshPivotTables();

                // Ensure the output directory exists
                string outputPath = "PivotTableDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}