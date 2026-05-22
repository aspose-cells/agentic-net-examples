using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;

namespace AsposeCellsDemo
{
    public class RemoveSlicerDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for a table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("A");
                worksheet.Cells["B4"].PutValue(30);

                // Add a table covering the data range
                int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Add a pivot table based on the same range
                int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "Pivot1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Add a slicer linked to the pivot table
                int slicerIndex = worksheet.Slicers.Add(pivotTable, "E3", "Category");
                Slicer slicer = worksheet.Slicers[slicerIndex];

                // ----- Removal logic -----
                // 1. Remove the slicer's connection to the pivot table (if any)
                slicer.RemovePivotConnection(pivotTable);

                // 2. Remove the slicer from the worksheet's slicer collection
                worksheet.Slicers.Remove(slicer);
                // -------------------------

                // Save the workbook to verify that the slicer has been removed
                workbook.Save("RemovedSlicer.xlsx");
                Console.WriteLine("Workbook saved successfully as 'RemovedSlicer.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveSlicerDemo.Run();
        }
    }
}