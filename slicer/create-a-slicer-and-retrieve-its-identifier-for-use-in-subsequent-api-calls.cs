using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Fruit");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            // Use the first field (Fruit) as a row field
            pivot.AddFieldToArea(PivotFieldType.Row, 0);

            // Access the slicer collection of the worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Add a slicer linked to the pivot table, placed at cell E1, filtering by the "Fruit" field
            int slicerIdx = slicers.Add(pivot, "E1", "Fruit");

            // Retrieve the slicer object using the returned index
            Slicer slicer = slicers[slicerIdx];

            // Optionally set a custom name for easier reference later
            slicer.Name = "FruitSlicer";

            // Output the slicer identifier (index) and name
            Console.WriteLine($"Slicer Index: {slicerIdx}");
            Console.WriteLine($"Slicer Name: {slicer.Name}");

            // Save the workbook (optional, demonstrates full lifecycle)
            workbook.Save("SlicerDemo.xlsx");
        }
    }
}