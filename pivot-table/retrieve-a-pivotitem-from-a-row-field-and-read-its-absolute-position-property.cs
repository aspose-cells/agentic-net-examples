// Title: Aspose.Cells for .NET: Get the absolute Position of a PivotItem in a row field (C#)
// Description: C# code that creates a workbook, adds sample data, builds a pivot table, puts the "Product" column in the Row area, refreshes the table, then accesses the first PivotItem of that row field and reads its Position property (the absolute index). The position is written to the console and the workbook is saved.
// Keywords: Aspose.Cells | C# | .NET | PivotItem.Position | pivot table row field | absolute position | retrieve pivot item index | Aspose.Cells PivotTable API | read pivot item position | PivotItem property
// Common Searches: Aspose.Cells get pivot item position C# | how to read PivotItem.Position in .NET | retrieve absolute index of row field item Aspose.Cells | pivot table item position property Aspose | C# Aspose.Cells PivotItem.Position example
// Developer Intent: Read the Position property of a specific PivotItem from a row field in an Aspose.Cells pivot table.
// Use Cases: Determine the exact order of row items for custom sorting logic. | Map pivot row items to external data structures using a stable index. | Debug pivot table layout by printing each PivotItem's absolute position.
// AI Prompts: Generate C# code that loops through all PivotItems in a row field of an Aspose.Cells pivot table and prints each item's Position. | Show how to manually reorder pivot row items by using the Position property of PivotItem in Aspose.Cells for .NET. | Explain the difference between PivotItem.Position and the displayed order when filters are applied in an Aspose.Cells pivot table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemPositionDemo
{
    // C# code that creates a workbook, adds sample data, builds a pivot table, puts the "Product" column in the Row area, refreshes the table, then accesses the first PivotItem of that row field and reads its Position property (the absolute index). The position is written to the console and the workbook is saved.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(3000);

            // Add a pivot table based on the data range
            int ptIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Add the "Product" column as a row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Refresh and calculate the pivot table to ensure items are generated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Retrieve the first row field
            PivotField rowField = pivotTable.RowFields[0];

            // Retrieve a specific PivotItem from the row field (e.g., the first item)
            PivotItem pivotItem = rowField.PivotItems[0];

            // Read the absolute Position property of the PivotItem
            int absolutePosition = pivotItem.Position;

            // Output the Position value
            Console.WriteLine("Absolute Position of the first PivotItem: " + absolutePosition);

            // Save the workbook (optional, demonstrates lifecycle rule)
            workbook.Save("PivotItemPositionDemo_out.xlsx");
        }
    }
}
