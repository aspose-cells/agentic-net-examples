using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
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

                // Add a new worksheet that will host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Determine the source data range (including headers)
                Aspose.Cells.Range sourceRange = sourceSheet.Cells.MaxDisplayRange;
                string sourceData = $"=SourceData!{sourceRange.Address}";

                // Add the pivot table to the new worksheet
                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceData, "A1", "SalesPivot");

                // Retrieve the created pivot table
                PivotTable pivotTable = pivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh the pivot table to calculate and display data
                pivotSheet.RefreshPivotTables();

                // Save the workbook
                string outputPath = "PivotTableDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            CreatePivotTable.Run();
        }
    }
}