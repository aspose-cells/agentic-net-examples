// Title: Add a slicer to a pivot table and obtain its index identifier with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to add a slicer linked to a pivot table, capture the returned slicer index, and retrieve the Slicer object via that index. | Assign a custom name to the retrieved slicer and save the workbook containing the slicer.
// Common Searches: Aspose.Cells C# get slicer index after adding slicer to a pivot table | retrieve slicer object by index in Aspose.Cells | how to name a slicer created with Aspose.Cells | reference slicer later using its identifier Aspose.Cells
// Tags: add slicer to pivot table Aspose.Cells | slicer index retrieval Aspose.Cells C# | set slicer name Aspose.Cells | save workbook with slicer Aspose.Cells | pivot table slicer identifier Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example creates a workbook, fills sample data, builds a pivot table, adds a slicer linked to that pivot table, captures the slicer's index, retrieves the slicer object using the index, assigns a custom name, and saves the file as SlicerDemo.xlsx.
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
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        // Use the first column (Fruit) as a row field
        pivot.AddFieldToArea(PivotFieldType.Row, 0);

        // Add a slicer linked to the pivot table
        // The method returns the slicer's index, which serves as its identifier
        int slicerIndex = sheet.Slicers.Add(pivot, "H1", "Fruit");

        // Retrieve the slicer object using the identifier
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Optionally set a custom name for later reference
        slicer.Name = "FruitSlicer";

        // Save the workbook to a file
        workbook.Save("SlicerDemo.xlsx");
    }
}
