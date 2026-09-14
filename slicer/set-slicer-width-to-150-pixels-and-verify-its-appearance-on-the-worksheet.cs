// Title: How to set a pivot table slicer width to 150 pixels and verify it using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a pivot table, inserts a slicer linked to a field, sets the slicer WidthPixel to 150, and saves the file with Aspose.Cells. | Show how to read the WidthPixel property of a slicer after assigning a value and output the result to the console in a .NET application. | Demonstrate verifying that the slicer size was applied correctly by comparing the expected pixel width with the actual property value.
// Common Searches: Aspose.Cells C# set slicer width to 150 pixels example | retrieve slicer WidthPixel value after setting it in Aspose.Cells | programmatically adjust Excel slicer size using Aspose.Cells for .NET | verify slicer dimensions in a workbook created with Aspose.Cells | how to change pivot table slicer pixel width in C#
// Tags: Aspose.Cells slicer WidthPixel property | C# set slicer pixel width | pivot table slicer sizing Aspose.Cells | verify slicer dimensions programmatically | Excel workbook slicer configuration .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

// The example creates a new workbook, fills it with sample data, builds a pivot table, adds a slicer linked to the 'Fruit' field, sets the slicer width to 150 pixels via the WidthPixel property, prints the width to the console for verification, and saves the workbook as SlicerWidthDemo.xlsx.
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
            sheet.Cells["A1"].PutValue("Fruit");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(200);

            // Add a pivot table based on the data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Add a slicer linked to the pivot table (correct parameter order: destination cell, then field name)
            int slicerIndex = sheet.Slicers.Add(pivot, "E5", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Set the slicer width to 150 pixels
            slicer.WidthPixel = 150;

            // Verify the width by outputting it to the console
            Console.WriteLine($"Slicer width (pixels): {slicer.WidthPixel}");

            // Save the workbook
            workbook.Save("SlicerWidthDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
