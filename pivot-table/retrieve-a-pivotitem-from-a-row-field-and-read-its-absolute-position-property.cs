// Title: Aspose.Cells for .NET (C#) – Get a PivotItem’s absolute Position from a row field
// Description: Creates a workbook, builds a pivot table, adds a row field, retrieves the first PivotItem, reads its Position (absolute index) and writes the value to the console before saving the file.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | PivotItem.Position | row field index | absolute position | retrieve pivot item | pivot item index example | Aspose.Cells tutorial
// Common Searches: Aspose.Cells get PivotItem position C# | how to read PivotItem absolute index in .NET | retrieve row field item position Aspose.Cells | PivotItem.Position property example | C# code to get pivot item order
// Developer Intent: Read the absolute Position of a specific PivotItem in a row field of an Aspose.Cells PivotTable.
// Use Cases: Custom sort row items based on their original order. | Validate the sequence of pivot items when generating automated reports. | Synchronize pivot item positions with external data structures or APIs.
// AI Prompts: Generate C# code with Aspose.Cells that accesses the first PivotItem of a row field and prints its Position. | Explain how the Position property of a PivotItem represents its absolute index within the row field hierarchy. | Show an example that loops through all PivotItems in a row field and outputs each item's Name and Position.

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;

// Creates a workbook, builds a pivot table, adds a row field, retrieves the first PivotItem, reads its Position (absolute index) and writes the value to the console before saving the file.
class RetrievePivotItemPosition
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("A");
        sheet.Cells["B4"].PutValue(150);
        sheet.Cells["A5"].PutValue("B");
        sheet.Cells["B5"].PutValue(250);

        // Add a pivot table to the worksheet
        int ptIndex = sheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];

        // Add a row field and a data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Refresh and calculate the pivot table
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Access the first row field
        PivotField rowField = pivotTable.RowFields[0];

        // Retrieve a specific PivotItem (e.g., the first item)
        PivotItem pivotItem = rowField.PivotItems[0];

        // Read its absolute Position property
        int absolutePosition = pivotItem.Position;

        // Display the result
        Console.WriteLine($"PivotItem '{pivotItem.Name}' absolute Position: {absolutePosition}");

        // Save the workbook
        workbook.Save("RetrievePivotItemPosition_out.xlsx");
    }
}
