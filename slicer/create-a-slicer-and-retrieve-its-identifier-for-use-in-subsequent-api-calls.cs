using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a pivot table
        sheet.Cells["A1"].PutValue("Fruit");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(15);

        // Add a pivot table based on the data range
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, 0); // Add "Fruit" as row field

        // Add a slicer linked to the pivot table
        SlicerCollection slicers = sheet.Slicers;
        int slicerIdx = slicers.Add(pivot, "H1", "Fruit"); // Destination cell H1, base field "Fruit"

        // Retrieve the slicer object using its identifier (index)
        Slicer slicer = slicers[slicerIdx];

        // Output identifier and name for verification
        Console.WriteLine("Slicer identifier (index): " + slicerIdx);
        Console.WriteLine("Slicer name: " + slicer.Name);

        // Save the workbook (optional)
        workbook.Save("SlicerDemo.xlsx");
    }
}