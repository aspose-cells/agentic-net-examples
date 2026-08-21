// Title: Check PivotTable RowField Position After Reordering with Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds a pivot table with two row fields, displays each field's Name and Position, moves the first row field to a new index using RowFields.Move, refreshes the pivot, and then prints the updated Position values to confirm they match the new order before saving the file.
// Keywords: Aspose.Cells pivot field position | RowFields.Move C# | verify pivot row field order | PivotTable Position property | Aspose.Cells .NET example | pivot table field reordering | check field Position after move | Aspose.Cells programming guide | C# pivot table manipulation | Excel pivot field verification
// Common Searches: Aspose.Cells how to read PivotField.Position | C# move pivot row field and verify order | RowFields.Move example Aspose.Cells | check pivot table field order after reordering | verify pivot field position .NET | Aspose.Cells pivot table field index
// Developer Intent: Ensure that the Position property of PivotField objects updates correctly after programmatically reordering row fields in a PivotTable.
// Use Cases: Automated testing of pivot table layout changes by comparing Position values before and after a move operation. | Dynamic report generation where row fields need to be reordered and their final positions must be validated. | Debugging and troubleshooting pivot table configurations in enterprise .NET applications.
// AI Prompts: Show a C# snippet using Aspose.Cells that moves a pivot row field and asserts the Position values before and after the move. | Explain how the PivotField.Position property is recalculated when RowFields.Move is called and how to retrieve it for validation. | Generate unit‑test code that verifies the correct ordering of PivotTable.RowFields after invoking the Move method.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This C# example creates a workbook, adds a pivot table with two row fields, displays each field's Name and Position, moves the first row field to a new index using RowFields.Move, refreshes the pivot, and then prints the updated Position values to confirm they match the new order before saving the file.
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

        // Add two fields to the Row area: "Category" (index 0) and "Value" (index 1)
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
        pivotTable.AddFieldToArea(PivotFieldType.Row, 1);

        // Display field names and their Position values before moving
        Console.WriteLine("Before move:");
        for (int i = 0; i < pivotTable.RowFields.Count; i++)
        {
            PivotField field = pivotTable.RowFields[i];
            Console.WriteLine($"Field {i}: {field.Name}, Position = {field.Position}");
        }

        // Move the first field (current position 0) to destination position 1
        pivotTable.RowFields.Move(0, 1);

        // Refresh and calculate the pivot table to apply the change
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Verify that Position values reflect the new order
        Console.WriteLine("\nAfter move:");
        for (int i = 0; i < pivotTable.RowFields.Count; i++)
        {
            PivotField field = pivotTable.RowFields[i];
            Console.WriteLine($"Field {i}: {field.Name}, Position = {field.Position}");
        }

        // Save the workbook
        workbook.Save("PivotFieldPositionVerification.xlsx");
    }
}
