using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotFieldShowItemsCountDemo
    {
        public static void Run()
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
                cells["A2"].PutValue("A");
                cells["B2"].PutValue(100);
                cells["A3"].PutValue("B");
                cells["B3"].PutValue(200);
                cells["A4"].PutValue("C");
                cells["B4"].PutValue(300);
                cells["A5"].PutValue("D");
                cells["B5"].PutValue(400);
                cells["A6"].PutValue("E");
                cells["B6"].PutValue(500);
                cells["A7"].PutValue("F");
                cells["B7"].PutValue(600);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B7", "D3", "DemoPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add the "Category" field to the row area and "Amount" to the data area
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // NOTE: The ShowItemsCount property is not available in this version of Aspose.Cells.
                // If needed, implement custom logic to filter displayed items.

                // Refresh and calculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotFieldShowItemsCountDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
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
            PivotFieldShowItemsCountDemo.Run();
        }
    }
}