// Title: Check Pivot Table Row Field Position After Reordering Using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a pivot table with two row fields, moves the first row field to a different index, refreshes the pivot, and prints each field's Name and Position before and after the move. | Generate a .NET example that demonstrates how to reorder pivot table row fields with Aspose.Cells and confirm the new Position values for each field.
// Common Searches: aspnet how to move a pivot table row field and read its Position property with Aspose.Cells | c# verify pivot field order after calling RowFields.Move in Aspose.Cells | display pivot table field positions after reordering rows using Aspose.Cells for .NET | check if pivot row field Position reflects new index after Move method in Aspose.Cells
// Tags: aspocells pivot rowfield reordering | c# check pivotfield position values | aspocells move row field index | pivot table field order validation .net | aspocells refreshdata calculate after move

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates creating a workbook, adding a pivot table with two row fields, moving the first row field to a new index, refreshing and calculating the pivot, and printing each field's Name and Position before and after the move, then saving the file.
class VerifyPivotFieldPositions
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add two fields to the Row area (Category and Value)
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0); // Category
        pivotTable.AddFieldToArea(PivotFieldType.Row, 1); // Value

        // Display field positions before moving
        Console.WriteLine("Before move:");
        for (int i = 0; i < pivotTable.RowFields.Count; i++)
        {
            PivotField field = pivotTable.RowFields[i];
            Console.WriteLine($"Index {i}: Name = {field.Name}, Position = {field.Position}");
        }

        // Move the first field (index 0) to the second position (index 1)
        pivotTable.RowFields.Move(0, 1);

        // Refresh and calculate the pivot table to apply the change
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Verify and display field positions after moving
        Console.WriteLine("\nAfter move:");
        for (int i = 0; i < pivotTable.RowFields.Count; i++)
        {
            PivotField field = pivotTable.RowFields[i];
            Console.WriteLine($"Index {i}: Name = {field.Name}, Position = {field.Position}");
        }

        // Save the workbook
        workbook.Save("PivotFieldPositionVerification.xlsx");
    }
}
