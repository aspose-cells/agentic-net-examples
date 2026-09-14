// Title: How to refresh a single PivotTable after modifying its source data with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that updates worksheet cells and then calls RefreshData and CalculateData on a specific PivotTable. | Show an example of creating a PivotTable, changing its source values, and programmatically refreshing only that PivotTable without affecting other tables. | Demonstrate how to save the workbook after refreshing a single PivotTable so the updated data appears in the output file.
// Common Searches: Aspose.Cells C# refresh only one pivot table after data edit | C# programmatically update pivot cache and recalculate specific pivot in Aspose.Cells | how to call RefreshData on a particular PivotTable using Aspose.Cells for .NET | RefreshData CalculateData example for single pivot table Aspose.Cells | update source cells and refresh a specific pivot table in a workbook with Aspose.Cells
// Tags: Aspose.Cells RefreshData method | single pivot table refresh .NET | pivot cache update Aspose.Cells | calculate pivot data C# | refresh specific pivot Aspose.Cells workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, fills it with source data, adds a PivotTable, performs an initial refresh, modifies the source cells, then calls RefreshData and CalculateData on that PivotTable only, and finally saves the workbook as RefreshedSpecificPivotTable.xlsx.
    public class RefreshSpecificPivotTableDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
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
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample source data for the pivot table
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(200);
            worksheet.Cells["A4"].PutValue("Apple");
            worksheet.Cells["B4"].PutValue(150);

            // Add a pivot table that uses the range A1:B4 as its data source
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Initial refresh and calculation so the pivot table has data
            pivotTable.RefreshData();      // Pull source data into the cache
            pivotTable.CalculateData();    // Recalculate the pivot report

            // Modify the underlying source data
            worksheet.Cells["B2"].PutValue(120);        // Updated sales for Apple
            worksheet.Cells["A3"].PutValue("Orange");  // Change product name to a new item

            // Refresh only this specific pivot table to reflect the changes
            pivotTable.RefreshData();      // Update the cache with new data
            pivotTable.CalculateData();    // Recalculate the pivot report

            // Save the workbook with the refreshed pivot table
            string outputPath = "RefreshedSpecificPivotTable.xlsx";
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                throw;
            }
        }
    }
}
