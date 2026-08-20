// Title: Aspose.Cells C# – Log Absolute Positions of All PivotItems per PivotField (Debug)
// Description: Creates a workbook, adds sample data, builds a pivot table, refreshes it, then iterates through Row, Column, and Page fields. For each PivotField it initializes the items, enumerates every PivotItem, and writes the item name with its absolute Position to the console for debugging, finally saving the workbook.
// Keywords: Aspose.Cells | C# | .NET | pivot table debugging | list pivot items | pivot item position | InitPivotItems | enumerate PivotFields | absolute position | Aspose.Cells API
// Common Searches: Aspose.Cells enumerate pivot items C# | Get pivot item position Aspose.Cells | Debug pivot table fields Aspose.Cells .NET | How to list PivotFields and PivotItems in Aspose.Cells | Retrieve absolute position of pivot items
// Developer Intent: The developer needs to walk through every PivotField in a pivot table, list all associated PivotItems, and output each item's absolute position for troubleshooting or validation.
// Use Cases: Validate that pivot items are generated correctly after a data refresh. | Detect missing or out‑of‑order items by comparing logged positions with expected layout. | Generate diagnostic logs for automated testing of pivot table structures.
// AI Prompts: Write C# code using Aspose.Cells to iterate all PivotFields and log each PivotItem's name and Position to a file. | Explain the role of InitPivotItems() in exposing PivotItems and when it should be invoked. | Show how to export pivot item names and their absolute positions to a CSV file with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDebug
{
    // Creates a workbook, adds sample data, builds a pivot table, refreshes it, then iterates through Row, Column, and Page fields. For each PivotField it initializes the items, enumerates every PivotItem, and writes the item name with its absolute Position to the console for debugging, finally saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue("Apple");
            sheet.Cells["C2"].PutValue(120);

            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["B3"].PutValue("Banana");
            sheet.Cells["C3"].PutValue(80);

            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B4"].PutValue("Carrot");
            sheet.Cells["C4"].PutValue(50);

            // Add a pivot table based on the data range
            int ptIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate to ensure items are generated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Helper method to process a collection of PivotFields
            void ProcessFields(PivotFieldCollection fields, string areaName)
            {
                foreach (PivotField field in fields)
                {
                    Console.WriteLine($"--- {areaName} Field: {field.Name} ---");
                    // Ensure pivot items are initialized
                    field.InitPivotItems();

                    foreach (PivotItem item in field.PivotItems)
                    {
                        // Log the item name and its absolute position
                        Console.WriteLine($"Item Name: {item.Name}, Position: {item.Position}");
                    }
                }
            }

            // Iterate over RowFields, ColumnFields, PageFields (if any)
            ProcessFields(pivotTable.RowFields, "Row");
            ProcessFields(pivotTable.ColumnFields, "Column");
            ProcessFields(pivotTable.PageFields, "Page");

            // Save the workbook (debug workbook can be inspected if needed)
            workbook.Save("PivotDebugOutput.xlsx");
        }
    }
}
