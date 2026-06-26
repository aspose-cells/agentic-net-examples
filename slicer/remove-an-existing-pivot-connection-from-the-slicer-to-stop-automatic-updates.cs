using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsExamples
{
    public class RemoveSlicerPivotConnectionDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["A4"].PutValue("Orange");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["B3"].PutValue(150);
            dataSheet.Cells["B4"].PutValue(200);

            // Add a worksheet to host the pivot table and slicer
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");

            // Create a pivot table based on the data range
            PivotTableCollection pivots = pivotSheet.PivotTables;
            int pivotIndex = pivots.Add("Data!A1:B4", "C3", "SalesPivot");
            PivotTable pivotTable = pivots[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sales field
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Add a slicer linked to the pivot table (field: Product)
            int slicerIndex = pivotSheet.Slicers.Add(pivotTable, "E3", "Product");
            Slicer slicer = pivotSheet.Slicers[slicerIndex];

            // Remove the pivot connection from the slicer to stop automatic updates
            slicer.RemovePivotConnection(pivotTable);

            // Define output file path
            string outputPath = "RemoveSlicerPivotConnection_out.xlsx";

            // Save the workbook (overwrite if exists)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
        }
    }
}