using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class PivotTableWithLightCells
{
    static void Main()
    {
        try
        {
            // Create a new workbook and set up source data worksheet
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Populate sample data
            Cells cells = sourceSheet.Cells;
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            string[] regions = { "North", "South", "East", "West" };
            string[] products = { "Apple", "Banana", "Cherry" };
            Random rnd = new Random();

            int row = 2;
            for (int i = 0; i < 20; i++)
            {
                cells[$"A{row}"].PutValue(regions[rnd.Next(regions.Length)]);
                cells[$"B{row}"].PutValue(products[rnd.Next(products.Length)]);
                cells[$"C{row}"].PutValue(rnd.Next(100, 1000));
                row++;
            }

            // Add worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Define source data range for the pivot cache
            string sourceData = $"=SourceData!{sourceSheet.Cells.MaxDisplayRange.Address}";

            // Add and configure the pivot table
            PivotTableCollection pivotTables = pivotSheet.PivotTables;
            int pivotIndex = pivotTables.Add(sourceData, "A1", "SalesPivot");
            PivotTable pivotTable = pivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivotTable.ShowInTabularForm();
            pivotTable.SaveData = true;

            // Refresh to reflect source data
            pivotSheet.RefreshPivotTables();

            // Define output path and ensure directory exists
            string outputPath = "PivotTableWithLightCells.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (standard save; LightCells API not required)
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}