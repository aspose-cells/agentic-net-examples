// Title: C# Example – Set PivotItem.Position After RefreshData & CalculateData with Aspose.Cells
// Description: Demonstrates how to refresh a pivot cache, calculate pivot data, and programmatically assign a custom order to each PivotItem by setting its Position property in Aspose.Cells for .NET. The workbook is saved as an Excel file.
// Keywords: Aspose.Cells | C# pivot table example | PivotItem.Position | RefreshData | CalculateData | reorder pivot items | Aspose.Cells for .NET | pivot table item ordering | Excel automation | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells set PivotItem.Position after RefreshData | C# reorder rows in Aspose.Cells pivot table | how to change pivot item order programmatically | RefreshData CalculateData Aspose.Cells example | move pivot items to specific positions in .NET
// Developer Intent: Programmatically define the exact sequence of row items in a pivot table by assigning a unique Position value to each PivotItem after the pivot cache has been refreshed and calculated.
// Use Cases: Apply a custom product display order after the source data is refreshed. | Maintain consistent row ordering across automated sales reports. | Synchronize pivot item positions with an external sorting list before exporting the workbook.
// AI Prompts: Generate C# code using Aspose.Cells to set PivotItem.Position for each row field after calling RefreshData and CalculateData. | Explain how the Position property influences pivot item ordering and how to assign unique positions when updating multiple items. | Show how to map a predefined list of product names to specific Position values after refreshing the pivot cache.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to refresh a pivot cache, calculate pivot data, and programmatically assign a custom order to each PivotItem by setting its Position property in Aspose.Cells for .NET. The workbook is saved as an Excel file.
    public class SetPivotItemPositionAfterRefresh
    {
        public static void Run()
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

                // Refresh the pivot cache data (using the available RefreshData method)
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // After refresh/calculation, set the Position property for each pivot item
                // Position specifies the absolute order of the item among all PivotItems
                PivotField rowField = pivotTable.RowFields[0];
                int targetPosition = 0; // Example: move every item to the first position sequentially
                foreach (PivotItem item in rowField.PivotItems)
                {
                    item.Position = targetPosition;
                    targetPosition++; // Increment to maintain unique positions
                }

                // Save the workbook
                workbook.Save("PivotItemPositionAfterRefresh.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetPivotItemPositionAfterRefresh.Run();
        }
    }
}
