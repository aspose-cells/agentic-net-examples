// Title: Detach a slicer from its linked pivot table in Aspose.Cells using C# to prevent auto‑refresh
// AI Prompts: Write C# code that creates a workbook, adds a pivot table and slicer, then uses Aspose.Cells to remove the slicer‑pivot connection. | Show how to call Slicer.RemovePivotConnection in Aspose.Cells for .NET to stop a slicer from updating automatically.
// Common Searches: Aspose.Cells C# how to break slicer link to pivot table | remove slicer pivot connection programmatically Aspose.Cells .NET | stop slicer auto refresh after pivot table changes using Aspose.Cells | C# example for Slicer.RemovePivotConnection in Aspose.Cells | disable slicer updates in Excel file with Aspose.Cells C#
// Tags: Aspose.Cells remove slicer‑pivot connection | C# detach slicer from pivot table | Aspose.Cells stop slicer auto refresh | Slicer.RemovePivotConnection method | Aspose.Cells workbook slicer management

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, adding sample data, building a pivot table, inserting a slicer linked to that pivot, removing the slicer‑pivot connection with RemovePivotConnection, and saving the workbook as an .xlsx file.
    public class RemoveSlicerPivotConnectionDemo
    {
        // Entry point for the example
        public static void Main()
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
            PivotTableCollection pivotTables = pivotSheet.PivotTables;
            int pivotIndex = pivotTables.Add("Data!A1:B4", "C3", "SalesPivot");
            PivotTable pivotTable = pivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales field

            // Refresh pivot cache and calculate data
            pivotTable.RefreshData();   // Correct API to refresh the cache
            pivotTable.CalculateData();

            // Add a slicer linked to the pivot table
            int slicerIndex = pivotSheet.Slicers.Add(pivotTable, "E3", "Product");
            Slicer slicer = pivotSheet.Slicers[slicerIndex];

            // Remove the pivot connection from the slicer to stop automatic updates
            slicer.RemovePivotConnection(pivotTable);

            // Save the workbook
            string outputPath = "RemoveSlicerPivotConnection_out.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
