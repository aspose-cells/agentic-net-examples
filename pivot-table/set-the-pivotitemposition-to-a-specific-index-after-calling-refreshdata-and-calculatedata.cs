// Title: Set PivotItem.Position to reorder rows after RefreshData and CalculateData with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that refreshes a pivot table, calculates its data, and then assigns each row field's PivotItem.Position to its zero‑based index using Aspose.Cells. | Show how to loop through a PivotField's PivotItems collection after calling RefreshData and CalculateData and set the Position property to reorder the pivot rows. | Explain why PivotItem.Position should be updated only after RefreshData and CalculateData when programmatically reordering pivot table rows in Aspose.Cells.
// Common Searches: Aspose.Cells C# set pivot item position after RefreshData and CalculateData | how to reorder pivot table rows programmatically using Aspose.Cells .NET | C# loop through PivotField.PivotItems to change Position property | update pivot item order after calculating pivot data with Aspose.Cells | set specific index for PivotItem.Position in Aspose.Cells pivot table
// Tags: pivotitem.position after refreshdata | reorder pivot table rows aspose.cells | iterate pivotitems set position c# | calculatedata before setting pivotitem position | aspnet workbook save pivot ordering

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemPositionDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table, refreshes and calculates the pivot cache, then iterates the row field's PivotItems to assign each item's Position property to its index, effectively reordering the rows before saving the file.
    class Program
    {
        static void Main()
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
                int ptIndex = sheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[ptIndex];

                // Add the "Product" field to the row area and "Sales" to the data area
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh the pivot cache and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Set the Position property for each pivot item.
                // Use a for‑loop to avoid modifying the collection while enumerating it.
                PivotField rowField = pivotTable.RowFields[0];
                for (int i = 0; i < rowField.PivotItems.Count; i++)
                {
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
