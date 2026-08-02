// Title: Aspose.Cells .NET: Set PivotItem.Position After RefreshData & CalculateData
// Description: Demonstrates how to create a workbook, build a pivot table, refresh its cache, calculate data, and then assign a custom Position value to each row‑field PivotItem using C# and Aspose.Cells, before saving the file.
// Keywords: Aspose.Cells C# pivot table | PivotItem.Position | RefreshData Aspose.Cells | CalculateData Aspose.Cells | reorder pivot rows .NET | programmatic pivot item ordering | Aspose.Cells tutorial | Excel automation C# | US developers Aspose.Cells | global Excel SDK
// Common Searches: how to change pivot item order with Aspose.Cells | set PivotItem.Position in C# after refreshing pivot cache | Aspose.Cells reorder row field items | update pivot item positions programmatically | Aspose.Cells pivot table custom sorting
// Developer Intent: Programmatically assign a specific Position index to each PivotItem after the pivot table cache has been refreshed and its data calculated.
// Use Cases: Maintain a predefined row order in reports that regenerate pivot caches daily. | Align pivot item sequence with an external ranking or priority list before exporting. | Ensure deterministic sorting of pivot rows when the source data changes.
// AI Prompts: Write C# code using Aspose.Cells to refresh a pivot table, calculate it, and set each PivotItem.Position to its loop index. | Show how to reorder pivot row items based on a custom list of positions with Aspose.Cells for .NET. | Provide an example that updates PivotItem.Position after RefreshData and CalculateData, then saves the workbook as an Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemPositionDemo
{
    // Demonstrates how to create a workbook, build a pivot table, refresh its cache, calculate data, and then assign a custom Position value to each row‑field PivotItem using C# and Aspose.Cells, before saving the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Orange");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["B4"].PutValue(1500);

                // Add a pivot table based on the data range
                int ptIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[ptIndex];

                // Add the "Product" field to the row area and "Sales" to the data area
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh the pivot cache and calculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Set the Position property for each pivot item.
                // Use a for‑loop to avoid modifying the collection while enumerating it.
                PivotField rowField = pivotTable.RowFields[0];
                for (int i = 0; i < rowField.PivotItems.Count; i++)
                {
                    // Example: assign each item's position to its index.
                    // Adjust the value as needed for your scenario.
                    rowField.PivotItems[i].Position = i;
                }

                // Save the workbook to a file
                workbook.Save("PivotItemPositionAfterRefresh.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
